using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;

namespace BatchInsertE2E
{
    internal static class Program
    {
        private static int Main(string[] args)
        {
            var options = ParseArgs(args);
            if (options == null)
            {
                Console.Error.WriteLine("Usage: BatchInsertE2E -CaseId AC-B03 -RepoRoot <path> [-Visible true]");
                Console.Error.WriteLine("   or: BatchInsertE2E -CaseIds AC-B01;AC-B02 -RepoRoot <path> [-Visible false]");
                return 2;
            }

            var caseIds = options.ResolveCaseIds();
            var caseResults = new List<Dictionary<string, object>>();
            Application word = null;

            try
            {
                var assetsRoot = Path.Combine(options.RepoRoot, "automation", "assets");
                var templatePath = Path.Combine(assetsRoot, "table-template.docx");
                if (!File.Exists(templatePath))
                {
                    WriteFailure(caseIds, caseResults, "missing table-template.docx; run automation/scripts/generate-fixtures.ps1");
                    return 1;
                }

                var wordToolsPath = Path.Combine(options.RepoRoot, "WordTools", "bin", "Release", "WordTools.dll");
                if (!File.Exists(wordToolsPath))
                {
                    WriteFailure(caseIds, caseResults, "missing WordTools.dll; build WordTools Release first");
                    return 1;
                }

                var wordToolsAssembly = Assembly.LoadFrom(wordToolsPath);
                word = new Application { Visible = options.Visible };

                var allPass = true;
                foreach (var caseId in caseIds)
                {
                    Document doc = null;
                    var payload = new Dictionary<string, object>
                    {
                        ["case_id"] = caseId,
                        ["pass"] = false,
                    };

                    try
                    {
                        doc = word.Documents.Open(templatePath);
                        doc.Activate();
                        PrepareSelection(doc, caseId);

                        var services = BuildServices(wordToolsAssembly, caseId);
                        var request = BuildRequest(wordToolsAssembly, caseId, assetsRoot);
                        ExecuteOrchestrator(wordToolsAssembly, word, request, services);

                        MaybeSaveOutputDocument(doc, caseId, payload);
                        FillResultPayload(payload, doc, services, caseId, payload.TryGetValue("saved_docx_path", out var saved) ? saved as string : null);
                        payload["pass"] = true;
                    }
                    catch (Exception ex)
                    {
                        payload["error"] = ex.Message;
                        allPass = false;
                    }
                    finally
                    {
                        if (doc != null)
                        {
                            doc.Close(false);
                        }
                    }

                    if (!(bool)payload["pass"])
                    {
                        allPass = false;
                    }

                    caseResults.Add(payload);
                }

                WriteSessionJson(caseIds, caseResults, allPass);
                return allPass ? 0 : 1;
            }
            catch (Exception ex)
            {
                WriteFailure(caseIds, caseResults, ex.Message);
                return 1;
            }
            finally
            {
                if (word != null)
                {
                    word.Quit(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(word);
                }
            }
        }

        private static void WriteFailure(string[] caseIds, List<Dictionary<string, object>> caseResults, string error)
        {
            if (caseResults.Count == 0)
            {
                caseResults.Add(new Dictionary<string, object>
                {
                    ["case_id"] = caseIds.Length > 0 ? caseIds[0] : "",
                    ["pass"] = false,
                    ["error"] = error,
                });
            }

            WriteSessionJson(caseIds, caseResults, false);
        }

        private static void WriteSessionJson(string[] caseIds, List<Dictionary<string, object>> caseResults, bool allPass)
        {
            if (caseIds.Length == 1 && caseResults.Count == 1)
            {
                WriteJson(caseResults[0]);
                return;
            }

            WriteJson(new Dictionary<string, object>
            {
                ["batch"] = true,
                ["pass"] = allPass,
                ["cases"] = caseResults,
            });
        }

        private static void MaybeSaveOutputDocument(Document doc, string caseId, Dictionary<string, object> payload)
        {
            if (!string.Equals(
                Environment.GetEnvironmentVariable("WORDTOOLS_E2E_SAVE_DOCX"),
                "1",
                StringComparison.Ordinal))
            {
                return;
            }

            payload["saved_docx_path"] = SaveOutputDocument(doc, caseId);
        }

        private static Options ParseArgs(string[] args)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (var i = 0; i < args.Length; i++)
            {
                var key = args[i].TrimStart('-');
                if (i + 1 < args.Length && !args[i + 1].StartsWith("-"))
                {
                    map[key] = args[++i];
                }
            }

            if (!map.TryGetValue("RepoRoot", out var repoRoot) || string.IsNullOrWhiteSpace(repoRoot))
            {
                return null;
            }

            map.TryGetValue("CaseId", out var caseId);
            map.TryGetValue("CaseIds", out var caseIds);
            if (string.IsNullOrWhiteSpace(caseId) && string.IsNullOrWhiteSpace(caseIds))
            {
                return null;
            }

            var visible = true;
            if (map.TryGetValue("Visible", out var visibleText))
            {
                visible = !string.Equals(visibleText, "false", StringComparison.OrdinalIgnoreCase);
            }

            return new Options
            {
                CaseId = caseId,
                CaseIds = caseIds,
                RepoRoot = Path.GetFullPath(repoRoot),
                Visible = visible,
            };
        }

        private static void PrepareSelection(Document doc, string caseId)
        {
            if (doc.Tables.Count == 0)
            {
                return;
            }

            var table = doc.Tables[1];
            switch (caseId)
            {
                case "AC-B01":
                    var outsideTable = doc.Content;
                    outsideTable.Collapse(WdCollapseDirection.wdCollapseEnd);
                    outsideTable.Select();
                    return;
                case "AC-B02":
                    table.Cell(1, 2).Range.Select();
                    return;
                default:
                    table.Cell(1, 1).Range.Select();
                    return;
            }
        }

        private static object BuildServices(Assembly wordToolsAssembly, string caseId)
        {
            var servicesType = wordToolsAssembly.GetType("WordTools.Services.InsertPhotosExecutionServices", throwOnError: true);
            var createHeadless = servicesType.GetMethod("CreateHeadless", BindingFlags.Public | BindingFlags.Static);
            if (createHeadless == null)
            {
                throw new InvalidOperationException(
                    "WordTools.dll is outdated (missing InsertPhotosExecutionServices.CreateHeadless). Rebuild WordTools Release.");
            }

            var cancelAfter = caseId == "AC-B06" ? 2 : 0;
            return createHeadless.Invoke(null, new object[] { 10, cancelAfter });
        }

        private static object BuildRequest(Assembly wordToolsAssembly, string caseId, string assetsRoot)
        {
            var requestType = wordToolsAssembly.GetType("WordTools.Forms.InsertPhotosRequest", throwOnError: true);
            var modeType = wordToolsAssembly.GetType("WordTools.Forms.InsertPhotosRequestMode", throwOnError: true);
            var folderMode = Enum.Parse(modeType, "Folder");
            var selectedMode = Enum.Parse(modeType, "SelectedFiles");
            var request = Activator.CreateInstance(requestType);

            switch (caseId)
            {
                case "AC-B04":
                case "AC-B07":
                case "AC-B08":
                case "AC-B11":
                    ConfigureFolderRequest(request, requestType, folderMode, assetsRoot, caseId);
                    break;
                case "AC-B05":
                    ConfigureSelectedRequest(request, requestType, selectedMode, assetsRoot, caseId);
                    break;
                default:
                    ConfigureSelectedRequest(request, requestType, selectedMode, assetsRoot, caseId == "AC-B06" ? "AC-B03" : caseId);
                    break;
            }

            return request;
        }

        private static void ConfigureSelectedRequest(object request, Type requestType, object selectedMode, string assetsRoot, string profileCaseId)
        {
            void Set(string name, object value)
            {
                requestType.GetProperty(name)?.SetValue(request, value, null);
            }

            Set("Mode", selectedMode);
            Set("IncludeRootImages", true);
            Set("IncludeSubFolderImages", true);

            switch (profileCaseId)
            {
                case "AC-B05":
                    Set("MinHeight", 30f);
                    Set("NeedDescription", true);
                    Set("UseFileNameAsDescription", true);
                    Set("UseFolderNameAsDescription", false);
                    Set("NeedAutoNumbering", false);
                    Set("NumberAlignment", 2);
                    Set("NumberPosition", 1);
                    Set("SelectedFiles", new[]
                    {
                        Path.Combine(assetsRoot, "images", "single", "01.jpg"),
                    });
                    return;

                case "AC-B09":
                    Set("MinHeight", 30f);
                    Set("NeedDescription", false);
                    Set("UseFileNameAsDescription", false);
                    Set("UseFolderNameAsDescription", false);
                    Set("NeedAutoNumbering", false);
                    Set("NumberAlignment", 2);
                    Set("NumberPosition", 1);
                    break;

                case "AC-B10":
                    Set("MinHeight", 30f);
                    Set("NeedDescription", false);
                    Set("UseFileNameAsDescription", true);
                    Set("UseFolderNameAsDescription", false);
                    Set("NeedAutoNumbering", true);
                    Set("NumberAlignment", 2);
                    Set("NumberPosition", 2);
                    break;

                case "AC-B12":
                    Set("MinHeight", 30f);
                    Set("NeedDescription", true);
                    Set("UseFileNameAsDescription", false);
                    Set("UseFolderNameAsDescription", false);
                    Set("NeedAutoNumbering", true);
                    Set("NumberAlignment", 1);
                    Set("NumberPosition", 1);
                    break;

                default:
                    Set("MinHeight", 30f);
                    Set("NeedDescription", false);
                    Set("UseFileNameAsDescription", true);
                    Set("UseFolderNameAsDescription", false);
                    Set("NeedAutoNumbering", true);
                    Set("NumberAlignment", 2);
                    Set("NumberPosition", 1);
                    break;
            }

            var folder = Path.Combine(assetsRoot, "images", "selected-4");
            Set("SelectedFiles", new[]
            {
                Path.Combine(folder, "01.jpg"),
                Path.Combine(folder, "02.jpg"),
                Path.Combine(folder, "03.jpg"),
                Path.Combine(folder, "04.jpg"),
            });
        }

        private static void ConfigureFolderRequest(object request, Type requestType, object folderMode, string assetsRoot, string caseId)
        {
            void Set(string name, object value)
            {
                requestType.GetProperty(name)?.SetValue(request, value, null);
            }

            Set("Mode", folderMode);
            Set("FolderPath", Path.Combine(assetsRoot, "images", "folder-root"));
            Set("MinHeight", 30f);
            Set("NeedDescription", false);
            Set("NeedAutoNumbering", true);
            Set("NumberAlignment", 2);
            Set("NumberPosition", 1);

            switch (caseId)
            {
                case "AC-B07":
                    Set("IncludeRootImages", true);
                    Set("IncludeSubFolderImages", false);
                    Set("UseFileNameAsDescription", true);
                    Set("UseFolderNameAsDescription", false);
                    return;

                case "AC-B08":
                    Set("IncludeRootImages", false);
                    Set("IncludeSubFolderImages", true);
                    Set("UseFileNameAsDescription", true);
                    Set("UseFolderNameAsDescription", false);
                    return;

                case "AC-B11":
                    Set("IncludeRootImages", true);
                    Set("IncludeSubFolderImages", true);
                    Set("UseFileNameAsDescription", false);
                    Set("UseFolderNameAsDescription", true);
                    return;

                default:
                    Set("IncludeRootImages", true);
                    Set("IncludeSubFolderImages", true);
                    Set("UseFileNameAsDescription", true);
                    Set("UseFolderNameAsDescription", false);
                    return;
            }
        }

        private static void ExecuteOrchestrator(Assembly wordToolsAssembly, Application word, object request, object services)
        {
            var orchestratorType = wordToolsAssembly.GetType("WordTools.Services.InsertPhotosOrchestrator", throwOnError: true);
            var orchestrator = Activator.CreateInstance(orchestratorType, word);
            var servicesType = services.GetType();
            var execute = orchestratorType.GetMethod("Execute", new[] { request.GetType(), servicesType });
            if (execute == null)
            {
                throw new InvalidOperationException(
                    "WordTools.dll is outdated (missing Execute(request, services)). Rebuild WordTools Release.");
            }

            execute.Invoke(orchestrator, new[] { request, services });
        }

        private static string SaveOutputDocument(Document doc, string caseId)
        {
            var outputDir = Path.Combine(Path.GetTempPath(), "wordtools-batch-insert-e2e");
            Directory.CreateDirectory(outputDir);
            var outputPath = Path.Combine(outputDir, caseId + ".docx");
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }

            doc.SaveAs2(outputPath);
            return outputPath;
        }

        private static void FillResultPayload(
            Dictionary<string, object> payload,
            Document doc,
            object services,
            string caseId,
            string savedDocxPath)
        {
            var progressReporter = services.GetType().GetProperty("ProgressReporter")?.GetValue(services, null);
            var notificationService = services.GetType().GetProperty("NotificationService")?.GetValue(services, null);

            payload["inline_shape_count"] = doc.InlineShapes.Count;
            payload["success_count"] = GetIntProperty(progressReporter, "LastSuccessCount");
            payload["fail_count"] = GetIntProperty(progressReporter, "LastFailCount");
            payload["cancelled"] = GetBoolProperty(progressReporter, "IsCancelled");
            payload["saved_docx_path"] = savedDocxPath;
            payload["warnings"] = GetNotificationMessages(notificationService, "Warnings");
            payload["informations"] = GetNotificationMessages(notificationService, "Informations");
            payload["has_numbered_description"] = DocumentAnalyzer.HasNumberedDescription(doc);
            payload["has_subfolder_title"] = DocumentAnalyzer.HasSubfolderTitle(doc, "sub-a");
            payload["last_image_row_col2_text"] = DocumentAnalyzer.GetLastImageRowCol2Text(doc);
            payload["description_samples"] = DocumentAnalyzer.CollectDescriptionSamples(doc);
            payload["has_number_after_description"] = DocumentAnalyzer.HasNumberAfterDescription(doc);
            payload["has_center_aligned_numbered_description"] = DocumentAnalyzer.HasCenterAlignedNumberedDescription(doc);
            payload["has_folder_name_description"] = DocumentAnalyzer.HasFolderNameDescription(doc);
            payload["has_manual_description_rows"] = DocumentAnalyzer.HasManualDescriptionRows(doc);
        }

        private static int GetIntProperty(object target, string propertyName)
        {
            if (target == null)
            {
                return 0;
            }

            var value = target.GetType().GetProperty(propertyName)?.GetValue(target, null);
            return value is int i ? i : 0;
        }

        private static bool GetBoolProperty(object target, string propertyName)
        {
            if (target == null)
            {
                return false;
            }

            var value = target.GetType().GetProperty(propertyName)?.GetValue(target, null);
            return value is bool b && b;
        }

        private static string[] GetNotificationMessages(object notificationService, string listPropertyName)
        {
            if (notificationService == null)
            {
                return Array.Empty<string>();
            }

            var list = notificationService.GetType().GetProperty(listPropertyName)?.GetValue(notificationService, null) as System.Collections.IEnumerable;
            if (list == null)
            {
                return Array.Empty<string>();
            }

            var messages = new List<string>();
            foreach (var item in list)
            {
                var message = item?.GetType().GetProperty("Message")?.GetValue(item, null) as string;
                if (!string.IsNullOrWhiteSpace(message))
                {
                    messages.Add(message);
                }
            }

            return messages.ToArray();
        }

        private static void WriteJson(Dictionary<string, object> payload)
        {
            var json = SimpleJson.Serialize(payload);
            Console.OutputEncoding = Encoding.UTF8;
            Console.WriteLine(json);
        }

        private sealed class Options
        {
            public string CaseId { get; set; }
            public string CaseIds { get; set; }
            public string RepoRoot { get; set; }
            public bool Visible { get; set; }

            public string[] ResolveCaseIds()
            {
                if (!string.IsNullOrWhiteSpace(CaseIds))
                {
                    return CaseIds.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
                }

                return new[] { CaseId };
            }
        }
    }
}
