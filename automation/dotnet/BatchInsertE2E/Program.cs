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
                return 2;
            }

            Application word = null;
            Document doc = null;
            var payload = new Dictionary<string, object>
            {
                ["case_id"] = options.CaseId,
                ["pass"] = false,
            };

            try
            {
                var assetsRoot = Path.Combine(options.RepoRoot, "automation", "assets");
                var templatePath = Path.Combine(assetsRoot, "table-template.docx");
                if (!File.Exists(templatePath))
                {
                    payload["error"] = "missing table-template.docx; run automation/scripts/generate-fixtures.ps1";
                    WriteJson(payload);
                    return 1;
                }

                var wordToolsPath = Path.Combine(options.RepoRoot, "WordTools", "bin", "Release", "WordTools.dll");
                if (!File.Exists(wordToolsPath))
                {
                    payload["error"] = "missing WordTools.dll; build WordTools Release first";
                    WriteJson(payload);
                    return 1;
                }

                var wordToolsAssembly = Assembly.LoadFrom(wordToolsPath);
                word = new Application { Visible = options.Visible };
                doc = word.Documents.Open(templatePath);
                doc.Activate();
                PrepareSelection(doc, options.CaseId);

                var services = BuildServices(wordToolsAssembly, options.CaseId);
                var request = BuildRequest(wordToolsAssembly, options.CaseId, assetsRoot);
                ExecuteOrchestrator(wordToolsAssembly, word, request, services);

                var outputDoc = SaveOutputDocument(doc, options.CaseId);
                FillResultPayload(payload, doc, services, options.CaseId, outputDoc);
                payload["pass"] = true;
                WriteJson(payload);
                return 0;
            }
            catch (Exception ex)
            {
                payload["error"] = ex.Message;
                WriteJson(payload);
                return 1;
            }
            finally
            {
                if (doc != null)
                {
                    doc.Close(false);
                }

                if (word != null)
                {
                    word.Quit(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(word);
                }
            }
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

            if (!map.TryGetValue("CaseId", out var caseId) || string.IsNullOrWhiteSpace(caseId))
            {
                return null;
            }

            if (!map.TryGetValue("RepoRoot", out var repoRoot) || string.IsNullOrWhiteSpace(repoRoot))
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
                    SetRequest(request, requestType, folderMode, assetsRoot, caseId);
                    break;
                case "AC-B05":
                    SetRequest(request, requestType, selectedMode, assetsRoot, caseId);
                    break;
                default:
                    SetRequest(request, requestType, selectedMode, assetsRoot, "AC-B03");
                    break;
            }

            return request;
        }

        private static void SetRequest(object request, Type requestType, object mode, string assetsRoot, string caseId)
        {
            void Set(string name, object value)
            {
                requestType.GetProperty(name)?.SetValue(request, value, null);
            }

            Set("Mode", mode);
            Set("MinHeight", 30f);
            Set("NeedDescription", true);
            Set("UseFileNameAsDescription", true);
            Set("UseFolderNameAsDescription", false);
            Set("NeedAutoNumbering", caseId != "AC-B05");
            Set("NumberAlignment", 0);
            Set("NumberPosition", 0);

            if (caseId == "AC-B04")
            {
                Set("FolderPath", Path.Combine(assetsRoot, "images", "folder-root"));
                Set("IncludeRootImages", true);
                Set("IncludeSubFolderImages", true);
                return;
            }

            if (caseId == "AC-B05")
            {
                Set("SelectedFiles", new[]
                {
                    Path.Combine(assetsRoot, "images", "single", "01.jpg"),
                });
                return;
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
            public string RepoRoot { get; set; }
            public bool Visible { get; set; }
        }
    }
}
