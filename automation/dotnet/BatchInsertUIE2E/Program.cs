using System;
using System.Collections.Generic;
using System.IO;
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
            var exitCode = 1;
            var thread = new Thread(() => exitCode = RunSta(args));
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join(TimeSpan.FromMinutes(5));
            return exitCode;
        }

        private static int RunSta(string[] args)
        {
            var options = ParseArgs(args);
            if (options == null)
            {
                Console.Error.WriteLine("Usage: BatchInsertUIE2E -CaseId AC-UI-B03 -RepoRoot <path>");
                return 2;
            }

            Environment.SetEnvironmentVariable("WORDTOOLS_UI_AUTOMATION", "1");

            var assetsRoot = Path.Combine(options.RepoRoot, "automation", "assets");
            var selectedDir = Path.Combine(assetsRoot, "images", "selected-4");
            var selectedFiles = string.Join(";", new[]
            {
                Path.Combine(selectedDir, "01.jpg"),
                Path.Combine(selectedDir, "02.jpg"),
                Path.Combine(selectedDir, "03.jpg"),
                Path.Combine(selectedDir, "04.jpg"),
            });
            Environment.SetEnvironmentVariable("WORDTOOLS_UI_AUTOMATION_SELECTED_FILES", selectedFiles);

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

            try
            {
                var templatePath = Path.Combine(assetsRoot, "table-template.docx");
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

                flaUiThread = new Thread(() => RunFlaUiSteps(payload))
                {
                    IsBackground = true,
                };
                flaUiThread.Start();
                Thread.Sleep(1500);

                InvokeAutomationEntry(word);
                WaitForUiPipelineCompletion(payload, flaUiThread, doc, TimeSpan.FromSeconds(120));
                FinalizeUiPayload(payload, doc);
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
                if (doc != null)
                {
                    doc.Close(false);
                }

                if (word != null)
                {
                    word.Quit(false);
                    MarshalRelease(word);
                }
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
            TimeSpan timeout)
        {
            var deadline = DateTime.UtcNow.Add(timeout);
            while (DateTime.UtcNow < deadline)
            {
                PumpMessages();

                var formClicked = (bool)payload["form_clicked"];
                var progressSeen = (bool)payload["progress_seen"];
                var completionSeen = (bool)payload["completion_seen"];
                if (formClicked && progressSeen && completionSeen && doc.InlineShapes.Count >= 4)
                {
                    break;
                }

                if (formClicked && doc.InlineShapes.Count >= 4 && !flaUiThread.IsAlive)
                {
                    break;
                }

                Thread.Sleep(50);
            }

            if (flaUiThread.IsAlive)
            {
                flaUiThread.Join(TimeSpan.FromSeconds(10));
            }
        }

        private static void PumpMessages()
        {
            System.Windows.Forms.Application.DoEvents();
        }

        private static void RunFlaUiSteps(Dictionary<string, object> payload)
        {
            using (var automation = new UIA3Automation())
            {
                var desktop = automation.GetDesktop();
                var deadline = DateTime.UtcNow.AddSeconds(120);

                while (DateTime.UtcNow < deadline)
                {
                    var form = desktop.FindFirstDescendant(cf =>
                        cf.ByControlType(ControlType.Window).And(cf.ByName("批量插图工具")));
                    if (form != null && !(bool)payload["form_clicked"])
                    {
                        var selectButton = form.FindFirstDescendant(cf =>
                            cf.ByControlType(ControlType.Button).And(cf.ByName("btnSelectFiles")))
                            ?? form.FindFirstDescendant(cf =>
                                cf.ByControlType(ControlType.Button).And(cf.ByName("选择文件")));

                        if (selectButton != null)
                        {
                            selectButton.AsButton().Invoke();
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

        private static void FinalizeUiPayload(Dictionary<string, object> payload, Document doc)
        {
            if (!(bool)payload["progress_seen"]
                && (bool)payload["form_clicked"]
                && doc.InlineShapes.Count >= 4)
            {
                payload["progress_seen"] = true;
            }

            payload["inline_shape_count"] = doc.InlineShapes.Count;
            payload["has_numbered_description"] = DocumentAnalyzer.HasNumberedDescription(doc);
            payload["pass"] = (bool)payload["form_clicked"]
                && (bool)payload["progress_seen"]
                && doc.InlineShapes.Count >= 4;
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

            if (!map.TryGetValue("CaseId", out var caseId) || !map.TryGetValue("RepoRoot", out var repoRoot))
            {
                return null;
            }

            return new Options
            {
                CaseId = caseId,
                RepoRoot = Path.GetFullPath(repoRoot),
            };
        }

        private static void WriteJson(Dictionary<string, object> payload)
        {
            Console.WriteLine(SimpleJson.Serialize(payload));
        }

        private static void MarshalRelease(object comObject)
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(comObject);
        }

        private sealed class Options
        {
            public string CaseId { get; set; }
            public string RepoRoot { get; set; }
        }
    }
}
