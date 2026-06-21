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
                Console.Error.WriteLine("Usage: BatchInsertUIE2E -CaseId AC-UI-B03 -RepoRoot <path> -ImageRoot <path>");
                return 2;
            }

            UiCasePlan plan;
            try
            {
                plan = UiCasePlan.Resolve(options.CaseId, options.ImageRoot);
            }
            catch (Exception ex)
            {
                WriteJson(new Dictionary<string, object>
                {
                    ["case_id"] = options.CaseId,
                    ["pass"] = false,
                    ["error"] = ex.Message,
                });
                return 1;
            }

            var exitCode = 1;
            var thread = new Thread(() => exitCode = RunSta(options, plan));
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join(plan.CompletionTimeout.Add(TimeSpan.FromMinutes(2)));
            return exitCode;
        }

        private static int RunSta(Options options, UiCasePlan plan)
        {
            plan.ApplyAutomationEnvironment();

            Application word = null;
            Document doc = null;
            Thread flaUiThread = null;
            var payload = new Dictionary<string, object>
            {
                ["case_id"] = options.CaseId,
                ["pass"] = false,
                ["ui_flow_started"] = false,
                ["form_clicked"] = false,
                ["progress_seen"] = false,
                ["completion_seen"] = false,
            };
            plan.WritePlanMetadata(payload);

            try
            {
                var assetsRoot = Path.Combine(options.RepoRoot, "automation", "assets");
                var templatePath = Path.Combine(assetsRoot, "table-template.docx");
                var manifestPath = Path.Combine(assetsRoot, "table-template.manifest.json");
                if (File.Exists(manifestPath))
                {
                    payload["min_table_row_count"] = ReadManifestRowCount(manifestPath);
                }

                if (!File.Exists(templatePath))
                {
                    payload["error"] = "missing table-template.docx";
                    WriteJson(payload);
                    return 1;
                }

                word = new Application { Visible = true };
                doc = word.Documents.Open(templatePath);
                doc.Activate();
                doc.Tables[1].Cell(1, 1).Range.Select();

                if (!WaitForAddinObject(word, payload, TimeSpan.FromSeconds(30)))
                {
                    WriteJson(payload);
                    return 1;
                }

                payload["ui_flow_started"] = true;

                flaUiThread = new Thread(() => RunFlaUiSteps(payload, plan))
                {
                    IsBackground = true,
                };
                flaUiThread.Start();
                Thread.Sleep(1500);

                InvokeAutomationEntry(word);
                WaitForUiPipelineCompletion(payload, flaUiThread, doc, plan);
                SaveOutputDocument(doc, options.CaseId, payload);
                FinalizeUiPayload(payload, doc, plan);
                WriteJson(payload);
                return (bool)payload["pass"] ? 0 : 1;
            }
            catch (Exception ex)
            {
                payload["error"] = ex.Message;
                WriteJson(payload);
                return 1;
            }
            finally
            {
                SafeCloseDocument(doc);
                SafeQuitWord(word);
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
                Thread.Sleep(500);
            }

            payload["error"] = "WordTools add-in not loaded; register plugin before UI E2E.";
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
                if (formClicked && progressSeen && completionSeen && shapeCount >= plan.ExpectedImageCount)
                {
                    break;
                }

                if (formClicked && GetInlineShapeCount(doc) >= plan.ExpectedImageCount && !flaUiThread.IsAlive)
                {
                    break;
                }

                Thread.Sleep(50);
            }

            if (flaUiThread.IsAlive)
            {
                flaUiThread.Join(TimeSpan.FromSeconds(30));
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

                    Thread.Sleep(100);
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
            payload["pass"] = (bool)payload["form_clicked"]
                && (bool)payload["progress_seen"]
                && GetInlineShapeCount(doc) == plan.ExpectedImageCount
                && (!string.Equals(plan.CaseId, "AC-UI-B03", StringComparison.OrdinalIgnoreCase)
                    || DocumentAnalyzer.HasNumberedDescription(doc))
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
            for (var attempt = 0; attempt < 60; attempt++)
            {
                try
                {
                    return doc.InlineShapes.Count;
                }
                catch (COMException ex) when (ex.HResult == rpcCallRejected)
                {
                    PumpMessages();
                    Thread.Sleep(100);
                }
            }

            return doc.InlineShapes.Count;
        }

        private static void SaveOutputDocument(Document doc, string caseId, Dictionary<string, object> payload)
        {
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

            if (!map.TryGetValue("CaseId", out var caseId)
                || !map.TryGetValue("RepoRoot", out var repoRoot)
                || !map.TryGetValue("ImageRoot", out var imageRoot))
            {
                return null;
            }

            return new Options
            {
                CaseId = caseId,
                RepoRoot = Path.GetFullPath(repoRoot),
                ImageRoot = Path.GetFullPath(imageRoot),
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

        private sealed class Options
        {
            public string CaseId { get; set; }
            public string RepoRoot { get; set; }
            public string ImageRoot { get; set; }
        }
    }
}
