using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.UIA3;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;

namespace BatchInsertUIE2E
{
    internal static class Program
    {
        private const string ProgId = "WordTools.ThisAddIn";

        private static int Main(string[] args)
        {
            var options = ParseArgs(args);
            if (options == null)
            {
                Console.Error.WriteLine(
                    "Usage: BatchInsertUIE2E -CaseId AC-UI-B03 -RepoRoot <path> -ImageRoot <path>");
                Console.Error.WriteLine(
                    "   or: BatchInsertUIE2E -CaseIds AC-UI-B05;AC-UI-B07 -RepoRoot <path> -ImageRoot <path>");
                return 2;
            }

            var caseIds = options.ResolveCaseIds();
            var sessionResult = new SessionResult();
            var thread = new Thread(() => sessionResult.ExitCode = RunSession(options, caseIds, sessionResult));
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join(ComputeSessionTimeout(caseIds, options).Add(TimeSpan.FromMinutes(2)));

            if (caseIds.Length == 1)
            {
                WriteJson(sessionResult.Cases.Count > 0 ? sessionResult.Cases[0] : new Dictionary<string, object>
                {
                    ["case_id"] = caseIds[0],
                    ["pass"] = false,
                    ["error"] = "no case result produced",
                });
            }
            else
            {
                WriteJson(new Dictionary<string, object>
                {
                    ["batch"] = true,
                    ["pass"] = sessionResult.ExitCode == 0,
                    ["cases"] = sessionResult.Cases,
                });
            }

            return sessionResult.ExitCode;
        }

        private static TimeSpan ComputeSessionTimeout(string[] caseIds, Options options)
        {
            var total = TimeSpan.Zero;
            foreach (var caseId in caseIds)
            {
                try
                {
                    var plan = UiCasePlan.Resolve(caseId, options.ImageRoot, options.RepoRoot);
                    total = total.Add(plan.CompletionTimeout);
                }
                catch
                {
                    total = total.Add(TimeSpan.FromMinutes(5));
                }
            }

            return total.Add(TimeSpan.FromMinutes(2));
        }

        private static int RunSession(Options options, string[] caseIds, SessionResult sessionResult)
        {
            Application word = null;
            var allPass = true;

            try
            {
                var assetsRoot = Path.Combine(options.RepoRoot, "automation", "assets");
                var templatePath = Path.Combine(assetsRoot, "table-template.docx");
                var manifestPath = Path.Combine(assetsRoot, "table-template.manifest.json");
                int? minTableRowCount = null;
                if (File.Exists(manifestPath))
                {
                    minTableRowCount = ReadManifestRowCount(manifestPath);
                }

                if (!File.Exists(templatePath))
                {
                    sessionResult.Cases.Add(new Dictionary<string, object>
                    {
                        ["case_id"] = caseIds[0],
                        ["pass"] = false,
                        ["error"] = "missing table-template.docx",
                    });
                    return 1;
                }

                var configPath = Path.Combine(Path.GetTempPath(), "wordtools-ui-automation-session.json");
                Environment.SetEnvironmentVariable("WORDTOOLS_UI_AUTOMATION", "1");
                Environment.SetEnvironmentVariable("WORDTOOLS_UI_AUTOMATION_CONFIG_FILE", configPath);

                word = new Application { Visible = true };
                if (!WaitForAddinObject(word, new Dictionary<string, object>(), TimeSpan.FromSeconds(30)))
                {
                    sessionResult.Cases.Add(new Dictionary<string, object>
                    {
                        ["case_id"] = caseIds[0],
                        ["pass"] = false,
                        ["error"] = "WordTools add-in not loaded; register plugin before UI E2E.",
                    });
                    return 1;
                }

                for (var index = 0; index < caseIds.Length; index++)
                {
                    var caseId = caseIds[index];
                    UiCasePlan plan;
                    try
                    {
                        plan = UiCasePlan.Resolve(caseId, options.ImageRoot, options.RepoRoot);
                    }
                    catch (Exception ex)
                    {
                        var failed = new Dictionary<string, object>
                        {
                            ["case_id"] = caseId,
                            ["pass"] = false,
                            ["error"] = ex.Message,
                        };
                        sessionResult.Cases.Add(failed);
                        allPass = false;
                        continue;
                    }

                    var caseResult = options.UseDirect
                        ? RunSingleCaseDirect(word, plan, templatePath, minTableRowCount, configPath)
                        : RunSingleCaseFlaUi(word, plan, templatePath, minTableRowCount, configPath, index == 0);
                    sessionResult.Cases.Add(caseResult);
                    if (!(bool)caseResult["pass"])
                    {
                        allPass = false;
                    }
                }

                sessionResult.ExitCode = allPass ? 0 : 1;
                return sessionResult.ExitCode;
            }
            catch (Exception ex)
            {
                sessionResult.Cases.Add(new Dictionary<string, object>
                {
                    ["case_id"] = caseIds.Length > 0 ? caseIds[0] : "",
                    ["pass"] = false,
                    ["error"] = ex.Message,
                });
                sessionResult.ExitCode = 1;
                return 1;
            }
            finally
            {
                SafeQuitWord(word);
            }
        }

        private static Dictionary<string, object> RunSingleCaseDirect(
            Application word,
            UiCasePlan plan,
            string templatePath,
            int? minTableRowCount,
            string configPath)
        {
            Document doc = null;
            var payload = new Dictionary<string, object>
            {
                ["case_id"] = plan.CaseId,
                ["pass"] = false,
                ["execution_mode"] = "direct",
                ["ui_flow_started"] = true,
                ["form_clicked"] = false,
                ["progress_seen"] = false,
                ["completion_seen"] = false,
            };
            plan.WritePlanMetadata(payload);
            if (minTableRowCount.HasValue)
            {
                payload["min_table_row_count"] = minTableRowCount.Value;
            }

            try
            {
                doc = word.Documents.Open(templatePath);
                doc.Activate();
                doc.Tables[1].Cell(1, 1).Range.Select();

                plan.WriteAutomationConfigFile(configPath);
                InvokeAutomationDirectExecute(word);
                payload["form_clicked"] = true;
                WaitForDirectCompletion(doc, plan);
                payload["progress_seen"] = true;
                payload["completion_seen"] = true;
                MaybeSaveOutputDocument(doc, plan.CaseId, payload);
                FinalizeUiPayload(payload, doc, plan);
                return payload;
            }
            catch (Exception ex)
            {
                payload["error"] = ex.Message;
                return payload;
            }
            finally
            {
                SafeCloseDocument(doc);
            }
        }

        private static void WaitForDirectCompletion(Document doc, UiCasePlan plan)
        {
            var deadline = DateTime.UtcNow.Add(TimeSpan.FromMinutes(3));
            while (DateTime.UtcNow < deadline)
            {
                PumpMessages();
                var shapeCount = GetInlineShapeCount(doc);
                if (plan.ExpectZeroImages)
                {
                    return;
                }

                if (shapeCount >= plan.ExpectedImageCount)
                {
                    return;
                }

                Thread.Sleep(50);
            }
        }

        private static void InvokeAutomationDirectExecute(Application word)
        {
            if (!TryGetAddinObject(word, out dynamic target))
            {
                throw new InvalidOperationException("WordTools add-in object not found.");
            }

            target.Automation_ExecuteFromConfig();
        }

        private static Dictionary<string, object> RunSingleCaseFlaUi(
            Application word,
            UiCasePlan plan,
            string templatePath,
            int? minTableRowCount,
            string configPath,
            bool firstCase)
        {
            Document doc = null;
            Thread flaUiThread = null;
            var payload = new Dictionary<string, object>
            {
                ["case_id"] = plan.CaseId,
                ["pass"] = false,
                ["execution_mode"] = "flaui",
                ["ui_flow_started"] = true,
                ["form_clicked"] = false,
                ["progress_seen"] = false,
                ["completion_seen"] = false,
            };
            plan.WritePlanMetadata(payload);
            if (minTableRowCount.HasValue)
            {
                payload["min_table_row_count"] = minTableRowCount.Value;
            }

            try
            {
                doc = word.Documents.Open(templatePath);
                doc.Activate();
                doc.Tables[1].Cell(1, 1).Range.Select();

                plan.ApplyAutomationEnvironment();
                plan.WriteAutomationConfigFile(configPath);

                flaUiThread = new Thread(() => RunFlaUiSteps(payload, plan))
                {
                    IsBackground = true,
                };
                flaUiThread.Start();
                Thread.Sleep(firstCase ? 500 : 200);

                InvokeAutomationEntry(word);
                WaitForUiPipelineCompletion(payload, flaUiThread, doc, plan);
                MaybeSaveOutputDocument(doc, plan.CaseId, payload);
                FinalizeUiPayload(payload, doc, plan);
                return payload;
            }
            catch (Exception ex)
            {
                payload["error"] = ex.Message;
                return payload;
            }
            finally
            {
                if (flaUiThread != null && flaUiThread.IsAlive)
                {
                    flaUiThread.Join(TimeSpan.FromSeconds(5));
                }

                SafeCloseDocument(doc);
                DismissTransientDialogs();
            }
        }

        private static void DismissTransientDialogs()
        {
            try
            {
                using (var automation = new UIA3Automation())
                {
                    var desktop = automation.GetDesktop();
                    foreach (var title in new[] { "完成", "提示", "批量插图" })
                    {
                        var dialog = desktop.FindFirstDescendant(cf =>
                            cf.ByControlType(ControlType.Window).And(cf.ByName(title)));
                        if (dialog == null)
                        {
                            continue;
                        }

                        var ok = dialog.FindFirstDescendant(cf =>
                            cf.ByControlType(ControlType.Button).And(cf.ByName("确定")))
                            ?? dialog.FindFirstDescendant(cf =>
                                cf.ByControlType(ControlType.Button).And(cf.ByName("OK")));
                        ok?.AsButton().Invoke();
                    }
                }
            }
            catch
            {
            }
        }

        private static void SafeCloseDocument(Document doc)
        {
            if (doc == null)
            {
                return;
            }

            try
            {
                doc.Close(false);
            }
            catch (COMException)
            {
            }
        }

        private static void SafeQuitWord(Application word)
        {
            if (word == null)
            {
                return;
            }

            try
            {
                word.Quit(false);
                MarshalRelease(word);
            }
            catch (COMException)
            {
            }
        }

        private static bool WaitForAddinObject(Application word, Dictionary<string, object> payload, TimeSpan timeout)
        {
            var deadline = DateTime.UtcNow.Add(timeout);
            while (DateTime.UtcNow < deadline)
            {
                if (TryGetAddinObject(word, out _))
                {
                    return true;
                }

                PumpMessages();
                Thread.Sleep(200);
            }

            if (payload.Count > 0)
            {
                payload["error"] = "WordTools add-in not loaded; register plugin before UI E2E.";
            }

            return false;
        }

        private static void WaitForUiPipelineCompletion(
            Dictionary<string, object> payload,
            Thread flaUiThread,
            Document doc,
            UiCasePlan plan)
        {
            var deadline = DateTime.UtcNow.Add(plan.CompletionTimeout);
            while (DateTime.UtcNow < deadline)
            {
                PumpMessages();

                var formClicked = (bool)payload["form_clicked"];
                var progressSeen = (bool)payload["progress_seen"];
                var completionSeen = (bool)payload["completion_seen"];
                var shapeCount = GetInlineShapeCount(doc);

                if (plan.ExpectZeroImages && formClicked)
                {
                    break;
                }

                if (formClicked && progressSeen && completionSeen && shapeCount >= plan.ExpectedImageCount)
                {
                    break;
                }

                if (formClicked && shapeCount >= plan.ExpectedImageCount && !flaUiThread.IsAlive)
                {
                    break;
                }

                Thread.Sleep(50);
            }

            if (flaUiThread.IsAlive)
            {
                flaUiThread.Join(TimeSpan.FromSeconds(15));
            }
        }

        private static void PumpMessages()
        {
            System.Windows.Forms.Application.DoEvents();
        }

        private static void RunFlaUiSteps(Dictionary<string, object> payload, UiCasePlan plan)
        {
            using (var automation = new UIA3Automation())
            {
                var desktop = automation.GetDesktop();
                var deadline = DateTime.UtcNow.Add(plan.CompletionTimeout);

                while (DateTime.UtcNow < deadline)
                {
                    var form = desktop.FindFirstDescendant(cf =>
                        cf.ByControlType(ControlType.Window).And(cf.ByName("批量插图工具")));
                    if (form != null && !(bool)payload["form_clicked"])
                    {
                        AutomationElement actionButton = null;
                        if (plan.FormAction == UiFormAction.InsertFromFolder)
                        {
                            actionButton = form.FindFirstDescendant(cf =>
                                    cf.ByControlType(ControlType.Button).And(cf.ByName("btnInsertFromFolder")))
                                ?? form.FindFirstDescendant(cf =>
                                    cf.ByControlType(ControlType.Button).And(cf.ByName("插入文件夹")));
                        }
                        else
                        {
                            actionButton = form.FindFirstDescendant(cf =>
                                    cf.ByControlType(ControlType.Button).And(cf.ByName("btnSelectFiles")))
                                ?? form.FindFirstDescendant(cf =>
                                    cf.ByControlType(ControlType.Button).And(cf.ByName("选择文件")));
                        }

                        if (actionButton != null)
                        {
                            actionButton.AsButton().Invoke();
                            payload["form_clicked"] = true;
                        }
                    }

                    var progress = desktop.FindFirstDescendant(cf =>
                            cf.ByControlType(ControlType.Window).And(cf.ByName("插入图片进度")))
                        ?? desktop.FindFirstDescendant(cf =>
                            cf.ByControlType(ControlType.Window).And(cf.ByName("ProgressForm")));
                    if (progress != null)
                    {
                        payload["progress_seen"] = true;
                    }

                    var batchConfirm = desktop.FindFirstDescendant(cf =>
                        cf.ByControlType(ControlType.Window).And(cf.ByName("批量插图")));
                    if (batchConfirm != null)
                    {
                        var okConfirm = batchConfirm.FindFirstDescendant(cf =>
                            cf.ByControlType(ControlType.Button).And(cf.ByName("确定")))
                            ?? batchConfirm.FindFirstDescendant(cf =>
                                cf.ByControlType(ControlType.Button).And(cf.ByName("OK")));
                        okConfirm?.AsButton().Invoke();
                    }

                    var warning = desktop.FindFirstDescendant(cf =>
                        cf.ByControlType(ControlType.Window).And(cf.ByName("提示")));
                    if (warning != null)
                    {
                        var okWarning = warning.FindFirstDescendant(cf =>
                            cf.ByControlType(ControlType.Button).And(cf.ByName("确定")));
                        okWarning?.AsButton().Invoke();
                    }

                    var completion = desktop.FindFirstDescendant(cf =>
                        cf.ByControlType(ControlType.Window).And(cf.ByName("完成")));
                    if (completion != null)
                    {
                        payload["completion_seen"] = true;
                        var ok = completion.FindFirstDescendant(cf =>
                            cf.ByControlType(ControlType.Button).And(cf.ByName("确定")));
                        ok?.AsButton().Invoke();
                    }

                    if ((bool)payload["form_clicked"] && (bool)payload["progress_seen"] && (bool)payload["completion_seen"])
                    {
                        break;
                    }

                    Thread.Sleep(75);
                }
            }
        }

        private static void FinalizeUiPayload(Dictionary<string, object> payload, Document doc, UiCasePlan plan)
        {
            if (!(bool)payload["progress_seen"]
                && (bool)payload["form_clicked"]
                && GetInlineShapeCount(doc) >= plan.ExpectedImageCount)
            {
                payload["progress_seen"] = true;
            }

            payload["inline_shape_count"] = GetInlineShapeCount(doc);
            payload["table_row_count"] = DocumentAnalyzer.GetTableRowCount(doc);
            payload["has_numbered_description"] = DocumentAnalyzer.HasNumberedDescription(doc);
            payload["has_subfolder_title"] = DocumentAnalyzer.HasAnySubfolderTitle(doc, plan.SubfolderTitleHints);
            payload["last_image_row_col2_text"] = DocumentAnalyzer.GetLastImageRowCol2Text(doc);
            payload["has_number_after_description"] = DocumentAnalyzer.HasNumberAfterDescription(doc);
            payload["has_center_aligned_numbered_description"] = DocumentAnalyzer.HasCenterAlignedNumberedDescription(doc);
            payload["has_left_aligned_numbered_description"] = DocumentAnalyzer.HasLeftAlignedNumberedDescription(doc);
            payload["has_folder_name_description"] = DocumentAnalyzer.HasFolderNameDescription(doc, plan.SubfolderTitleHints);
            payload["has_manual_description_rows"] = DocumentAnalyzer.HasManualDescriptionRows(doc);

            var shapeCount = GetInlineShapeCount(doc);
            if (plan.ExpectZeroImages)
            {
                payload["pass"] = (bool)payload["form_clicked"] && shapeCount == 0;
                return;
            }

            payload["pass"] = (bool)payload["form_clicked"]
                && (bool)payload["progress_seen"]
                && shapeCount == plan.ExpectedImageCount
                && MeetsMinimumTableRowCount(payload, doc);
        }

        private static bool MeetsMinimumTableRowCount(Dictionary<string, object> payload, Document doc)
        {
            if (!payload.TryGetValue("min_table_row_count", out var minObj))
            {
                return true;
            }

            var minRows = Convert.ToInt32(minObj);
            return DocumentAnalyzer.GetTableRowCount(doc) >= minRows;
        }

        private static int GetInlineShapeCount(Document doc)
        {
            const int rpcCallRejected = unchecked((int)0x80010001);
            for (var attempt = 0; attempt < 30; attempt++)
            {
                try
                {
                    return doc.InlineShapes.Count;
                }
                catch (COMException ex) when (ex.HResult == rpcCallRejected)
                {
                    PumpMessages();
                    Thread.Sleep(50);
                }
            }

            return doc.InlineShapes.Count;
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

            if (doc == null)
            {
                return;
            }

            try
            {
                var outputDir = Path.Combine(Path.GetTempPath(), "wordtools-batch-insert-ui-e2e");
                Directory.CreateDirectory(outputDir);
                var outputPath = Path.Combine(outputDir, caseId + ".docx");
                if (File.Exists(outputPath))
                {
                    File.Delete(outputPath);
                }

                doc.SaveAs2(outputPath);
                payload["saved_docx_path"] = outputPath;
            }
            catch (Exception ex)
            {
                payload["save_error"] = ex.Message;
            }
        }

        private static void InvokeAutomationEntry(Application word)
        {
            if (!TryGetAddinObject(word, out dynamic target))
            {
                throw new InvalidOperationException("WordTools add-in object not found.");
            }

            target.Automation_ShowInsertPhotosForm();
        }

        private static bool TryGetAddinObject(Application word, out dynamic target)
        {
            target = null;
            dynamic dWord = word;
            dynamic addins = dWord.COMAddIns;
            int count = addins.Count;
            for (var i = 1; i <= count; i++)
            {
                dynamic addin = addins.Item(i);
                string progId = addin.ProgId;
                if (!string.Equals(progId, ProgId, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (!addin.Connect)
                {
                    addin.Connect = true;
                }

                target = addin.Object;
                if (target != null)
                {
                    return true;
                }
            }

            return false;
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

            if (!map.TryGetValue("RepoRoot", out var repoRoot)
                || !map.TryGetValue("ImageRoot", out var imageRoot))
            {
                return null;
            }

            map.TryGetValue("CaseId", out var caseId);
            map.TryGetValue("CaseIds", out var caseIds);

            if (string.IsNullOrWhiteSpace(caseId) && string.IsNullOrWhiteSpace(caseIds))
            {
                return null;
            }

            var useDirect = true;
            if (map.TryGetValue("Direct", out var directText))
            {
                useDirect = !string.Equals(directText, "false", StringComparison.OrdinalIgnoreCase);
            }

            return new Options
            {
                CaseId = caseId,
                CaseIds = caseIds,
                RepoRoot = Path.GetFullPath(repoRoot),
                ImageRoot = Path.GetFullPath(imageRoot),
                UseDirect = useDirect,
            };
        }

        private static void WriteJson(Dictionary<string, object> payload)
        {
            Console.WriteLine(SimpleJson.Serialize(payload));
        }

        private static int ReadManifestRowCount(string manifestPath)
        {
            try
            {
                var json = File.ReadAllText(manifestPath);
                var marker = "\"row_count\"";
                var index = json.IndexOf(marker, StringComparison.Ordinal);
                if (index < 0)
                {
                    return 8;
                }

                var colon = json.IndexOf(':', index);
                var end = json.IndexOfAny(new[] { ',', '}' }, colon + 1);
                if (colon < 0 || end < 0)
                {
                    return 8;
                }

                var raw = json.Substring(colon + 1, end - colon - 1).Trim();
                return int.TryParse(raw, out var count) ? count : 8;
            }
            catch
            {
                return 8;
            }
        }

        private static void MarshalRelease(object comObject)
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(comObject);
        }

        private sealed class SessionResult
        {
            public int ExitCode { get; set; } = 1;
            public List<Dictionary<string, object>> Cases { get; } = new List<Dictionary<string, object>>();
        }

        private sealed class Options
        {
            public string CaseId { get; set; }
            public string CaseIds { get; set; }
            public string RepoRoot { get; set; }
            public string ImageRoot { get; set; }
            public bool UseDirect { get; set; } = true;

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
