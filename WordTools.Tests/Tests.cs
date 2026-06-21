using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using WordTools.Forms;
using WordTools.Services;
using Xunit;

namespace WordTools.Tests
{
    public class Tests
    {
        [Fact]
        public void TestProgressButtonSwitchesToCloseAfterCompletion()
        {
            var controller = new ProgressFormStateController();

            ProgressButtonAction firstClick = controller.HandleButtonClick();
            Assert.Equal(ProgressButtonAction.CancelRequested, firstClick);
            Assert.True(controller.IsCancelled, "controller should record cancellation request");

            controller.MarkCompleted();

            ProgressButtonAction completionClick = controller.HandleButtonClick();
            Assert.Equal(ProgressButtonAction.CloseRequested, completionClick);
            Assert.Equal("关闭", controller.ButtonText);
            Assert.True(controller.IsButtonEnabled, "completed state should re-enable the primary button");
        }

        [Fact]
        public void TestBenchmarkLogWritesHeaderOnce()
        {
            string tempFile = Path.Combine(Path.GetTempPath(), "wordtools-benchmark-" + Guid.NewGuid().ToString("N") + ".csv");

            try
            {
                var entry = new BenchmarkLogEntry
                {
                    RunMode = "SelectedFiles",
                    Status = "Completed",
                    TotalFiles = 2,
                    SuccessCount = 2,
                    FailCount = 0
                };

                BenchmarkLogService.AppendCsv(tempFile, entry);
                BenchmarkLogService.AppendCsv(tempFile, entry);

                string[] lines = File.ReadAllLines(tempFile);
                Assert.Equal(3, lines.Length);
                Assert.Contains("timestamp_utc", lines[0]);
            }
            finally
            {
                if (File.Exists(tempFile))
                {
                    File.Delete(tempFile);
                }
            }
        }

        [Fact]
        public void TestBenchmarkLogEscapesCsvFields()
        {
            string tempFile = Path.Combine(Path.GetTempPath(), "wordtools-benchmark-" + Guid.NewGuid().ToString("N") + ".csv");

            try
            {
                var entry = new BenchmarkLogEntry
                {
                    RunMode = "Folder",
                    Status = "Completed",
                    SourcePath = "D:\\Data,Set\\Input",
                    ErrorMessage = "line1\r\nline2 \"quoted\""
                };

                BenchmarkLogService.AppendCsv(tempFile, entry);

                string[] lines = File.ReadAllLines(tempFile);
                Assert.Contains("\"D:\\Data,Set\\Input\"", lines[1]);
                Assert.Contains("\"line1 line2 \"\"quoted\"\"\"", lines[1]);
            }
            finally
            {
                if (File.Exists(tempFile))
                {
                    File.Delete(tempFile);
                }
            }
        }

        [Fact]
        public void TestLoggingOptionsDisableBenchmarkWithoutDetailedLog()
        {
            LoggingOptionsState state = LoggingOptionsStateController.Normalize(false, true);

            Assert.False(state.DetailedLoggingEnabled, "detailed logging should stay disabled");
            Assert.False(state.BenchmarkLoggingEnabled, "benchmark logging should be cleared when detailed logging is disabled");
            Assert.False(LoggingOptionsStateController.ShouldWriteBenchmarkLog(false, true),
                "benchmark log should not write when detailed logging is disabled");
            Assert.False(LoggingOptionsStateController.ShouldShowDetailedLog(false),
                "diagnostic detail should stay hidden by default");
        }

        [Fact]
        public void TestLoggingOptionsAllowBenchmarkWithDetailedLog()
        {
            LoggingOptionsState state = LoggingOptionsStateController.Normalize(true, true);

            Assert.True(state.DetailedLoggingEnabled, "detailed logging should remain enabled");
            Assert.True(state.BenchmarkLoggingEnabled, "benchmark logging should remain enabled when detailed logging is enabled");
            Assert.True(LoggingOptionsStateController.ShouldWriteBenchmarkLog(true, true),
                "benchmark log should write only when both toggles are enabled");
            Assert.True(LoggingOptionsStateController.ShouldShowDetailedLog(true),
                "diagnostic detail should appear when detailed logging is enabled");
        }

        [Fact]
        public void TestInsertionDiagnosticsReportHotPathBreakdown()
        {
            var diagnostics = new InsertionPerformanceDiagnostics();
            diagnostics.RecordCellAvailabilityCheck(12);
            diagnostics.RecordCellAvailabilityCheck(8);
            diagnostics.RecordFloatingShapeLookup(5);
            diagnostics.RecordOverwriteClear(7);
            diagnostics.RecordAddPicture(42);
            diagnostics.RecordDescriptionWrite(11);

            string detail = diagnostics.BuildDetailedLog();

            Assert.Contains("行可用性检查: 20ms (2次)", detail);
            Assert.Contains("浮动图片探测: 5ms (1次)", detail);
            Assert.Contains("覆盖清理: 7ms (1次)", detail);
            Assert.Contains("AddPicture: 42ms (1次)", detail);
            Assert.Contains("描述/编号写入: 11ms (1次)", detail);
        }

        [Fact]
        public void TestInsertionDiagnosticsReportExtendedHotPathBreakdown()
        {
            var diagnostics = new InsertionPerformanceDiagnostics();
            diagnostics.RecordCellValidation(9);
            diagnostics.RecordPictureSizing(14);
            diagnostics.RecordProgressUi(6);

            string detail = diagnostics.BuildDetailedLog();

            Assert.Contains("Cell validation: 9ms (1", detail);
            Assert.Contains("Picture sizing: 14ms (1", detail);
            Assert.Contains("Progress UI: 6ms (1", detail);
        }

        [Fact]
        public void TestFloatingShapeIndexFindsAnchorsWithinCellRange()
        {
            FloatingShapeIndex index = FloatingShapeIndex.Create(new[]
            {
                new FloatingShapeAnchor(10, 18),
                new FloatingShapeAnchor(30, 38),
                new FloatingShapeAnchor(60, 70)
            });

            Assert.True(index.HasShapeInRange(1, 20),
                "index should find anchors fully contained in the current cell range");
            Assert.True(index.HasShapeInRange(25, 40),
                "index should find later anchors in a different cell range");
            Assert.False(index.HasShapeInRange(19, 29),
                "index should not report anchors outside the current cell range");
            Assert.False(index.HasShapeInRange(61, 65),
                "index should require the whole anchor to fit in the cell range");
        }

        [Fact]
        public void TestBatchContextReusesFloatingShapeIndexUntilInvalidated()
        {
            var context = new ImageInsertionBatchContext();
            int factoryCalls = 0;

            FloatingShapeIndex first = context.GetOrCreateFloatingShapeIndex(() =>
            {
                factoryCalls++;
                return FloatingShapeIndex.Create(new[] { new FloatingShapeAnchor(5, 9) });
            });

            FloatingShapeIndex second = context.GetOrCreateFloatingShapeIndex(() =>
            {
                factoryCalls++;
                return FloatingShapeIndex.Create(new[] { new FloatingShapeAnchor(20, 29) });
            });

            Assert.Equal(1, factoryCalls);
            Assert.Same(first, second);

            context.InvalidateFloatingShapeIndex();

            FloatingShapeIndex third = context.GetOrCreateFloatingShapeIndex(() =>
            {
                factoryCalls++;
                return FloatingShapeIndex.Create(new[] { new FloatingShapeAnchor(20, 29) });
            });

            Assert.Equal(2, factoryCalls);
            Assert.NotSame(first, third);
        }

        [Fact]
        public void TestBatchContextCachesRowAvailabilityUntilCleared()
        {
            var context = new ImageInsertionBatchContext();
            var row = new ImageRowAvailability(88, ImageCellAvailability.Available, ImageCellAvailability.OverwriteText);

            Assert.False(context.TryGetCachedRowAvailability(88, out _),
                "empty context should not report a cached row");

            context.CacheRowAvailability(row);

            Assert.True(context.TryGetCachedRowAvailability(88, out var cachedRow),
                "cached row should be returned for later lookups");
            Assert.Equal(ImageCellAvailability.OverwriteText, cachedRow.RightCell);

            context.ClearRowAvailability();

            Assert.False(context.TryGetCachedRowAvailability(88, out _),
                "clearing the cache should remove stored row availability");
        }

        [Fact]
        public void TestSummaryRequiredForWarningsWithoutFailures()
        {
            Assert.True(InsertionSummaryFormatter.ShouldShowSummary(
                    0,
                    new List<int> { 3 },
                    new List<string>()),
                "merged-cell warnings alone should still trigger the unified summary");

            Assert.True(InsertionSummaryFormatter.ShouldShowSummary(
                    0,
                    new List<int>(),
                    new List<string> { "第7行第1列已有图片，已覆盖插入新图片" }),
                "overwrite warnings alone should still trigger the unified summary");
        }

        [Fact]
        public void TestSummaryMessageIncludesMergedAndOverwriteSections()
        {
            string message = InsertionSummaryFormatter.BuildSummaryMessage(
                5,
                1,
                "1.2秒",
                "",
                new List<(string fileName, string errorReason)>
                {
                    ("a.jpg", "文件不存在或已被移动")
                },
                new List<int> { 8 },
                new List<string> { "第7行第1列已有文本，已覆盖插入新图片" });

            Assert.Contains("失败: 1 张", message);
            Assert.Contains("合并单元格绕开: 1 处", message);
            Assert.Contains("覆盖插图提示: 1 处", message);
            Assert.Contains("第8行", message);
            Assert.Contains("第7行第1列已有文本", message);
        }

        [Fact]
        public void TestSummaryMessageExplainsDetailsAction()
        {
            string message = InsertionSummaryFormatter.BuildDetailsPrompt();

            Assert.Contains("是(Y)", message);
            Assert.Contains("查看详情", message);
            Assert.Contains("否(N)", message);
        }

        [Fact]
        public void TestFailureDetailsButtonsUseExplicitLabels()
        {
            using var form = new FailureDetailsForm(
                new List<(string fileName, string errorReason)> { ("a.jpg", "文件不存在或已被移动") },
                new List<int> { 11 },
                new List<string> { "第7行第1列已有文本，已覆盖插入新图片" });

            var buttons = GetAllControls(form).OfType<Button>().ToList();
            var buttonTexts = buttons.Select(button => button.Text).ToList();

            Assert.Contains("复制详情", buttonTexts);
            Assert.Contains("关闭窗口", buttonTexts);

            Button copyButton = buttons.Single(button => button.Text == "复制详情");
            Button closeButton = buttons.Single(button => button.Text == "关闭窗口");
            Assert.True(copyButton.Width >= 96, "copy-details button should be wide enough to avoid clipping");
            Assert.True(closeButton.Width >= 96, "close-window button should be wide enough to avoid clipping");
            Assert.True(copyButton.Height >= 32, "copy-details button should be tall enough to avoid clipping");
            Assert.True(closeButton.Height >= 32, "close-window button should be tall enough to avoid clipping");
        }

        [Fact]
        public void TestProgressServiceSourceHasNoMojibake()
        {
            string sourcePath = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "WordTools", "Services", "ProgressService.cs");
            string source = File.ReadAllText(Path.GetFullPath(sourcePath));
            string[] mojibakeMarkers =
            {
                "\u934f\u5825\u20ac\u20ac",
                "\u6769\u6d98\u5bb3\u7ed0\u6940\u5f5b",
                "\u93c2\u56e6\u6b22\u93b5\u5f52\u567a",
                "\u93c8\ue046\u7161\u95bf\u6b12\ue1e4",
                "\u935a\u581d\u82df\u9357\u66de\u5393\u93cd",
                "\u7487\ufe3d\u510f"
            };

            foreach (string marker in mojibakeMarkers)
            {
                Assert.DoesNotContain(marker, source);
            }
        }

        [Fact]
        public void TestInsertPhotosFormSourceHasNoLoggingControls()
        {
            string sourcePath = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "WordTools", "Forms", "InsertPhotosForm.cs");
            string source = File.ReadAllText(Path.GetFullPath(sourcePath));

            Assert.DoesNotContain("chkDetailedLogging", source);
            Assert.DoesNotContain("chkBenchmarkLogging", source);
            Assert.DoesNotContain("CreateLoggingOptionsSection", source);
        }

        [Fact]
        public void TestInsertPhotosFormClosesBeforeLaunchingInsertionWork()
        {
            string sourcePath = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "WordTools", "Forms", "InsertPhotosForm.cs");
            string source = File.ReadAllText(Path.GetFullPath(sourcePath));
            string orchestratorPath = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "WordTools", "Services", "InsertPhotosOrchestrator.cs");
            string orchestratorSource = File.ReadAllText(Path.GetFullPath(orchestratorPath));
            string addInPath = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "WordTools", "ThisAddIn.cs");
            string addInSource = File.ReadAllText(Path.GetFullPath(addInPath));

            Assert.DoesNotContain("Hide();", source);
            Assert.Contains("PendingRequest = new InsertPhotosRequest", source);
            Assert.Contains("DialogResult = DialogResult.OK;", source);
            Assert.Contains("ExecuteDeferred(pendingRequest);", orchestratorSource);
            Assert.True(
                orchestratorSource.Contains("System.Windows.Forms.Timer") || orchestratorSource.Contains("new Timer"),
                "orchestrator should defer insertion until after the modal dialog has fully closed");
            string ribbonControllerPath = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "WordTools", "RibbonController.cs");
            string ribbonControllerSource = File.ReadAllText(Path.GetFullPath(ribbonControllerPath));
            Assert.Contains("InsertPhotosOrchestrator", ribbonControllerSource);
            Assert.Contains("_ribbonController.OnInsertPhotosClick", addInSource);
        }

        [Fact]
        public void ThisAddIn_ribbon_callbacks_delegate_to_ribbon_controller()
        {
            string addInSource = ReadProjectSource(Path.Combine("WordTools", "ThisAddIn.cs"));
            string ribbonControllerSource = ReadProjectSource(Path.Combine("WordTools", "RibbonController.cs"));

            Assert.Contains("_ribbonController.OnInsertPhotosClick", addInSource);
            Assert.Contains("_ribbonController.OnRefreshNumberingClick", addInSource);
            Assert.DoesNotContain("NumberingRefreshService", addInSource);
            Assert.Contains("InsertPhotosOrchestrator", ribbonControllerSource);
            Assert.Contains("new InsertPhotosOrchestrator", addInSource);
        }

        [Fact]
        public void Numbering_refresh_entry_owned_by_ribbon_controller()
        {
            string ribbonControllerSource = ReadProjectSource(Path.Combine("WordTools", "RibbonController.cs"));

            Assert.Contains("NumberingRefreshService", ribbonControllerSource);
            Assert.Contains("RefreshFromCurrentSelection", ribbonControllerSource);
        }

        [Fact]
        public void InsertPhotosOrchestrator_does_not_call_message_box_directly()
        {
            string source = ReadProjectSource(Path.Combine("WordTools", "Services", "InsertPhotosOrchestrator.cs"));
            Assert.DoesNotContain("MessageBox.Show", source);
        }

        [Fact]
        public void RibbonController_does_not_call_message_box_directly()
        {
            string source = ReadProjectSource(Path.Combine("WordTools", "RibbonController.cs"));
            Assert.DoesNotContain("MessageBox.Show", source);
        }

        [Fact]
        public void InsertPhotosForm_does_not_call_message_box_directly()
        {
            string source = ReadProjectSource(Path.Combine("WordTools", "Forms", "InsertPhotosForm.cs"));
            Assert.DoesNotContain("MessageBox.Show", source);
        }

        [Fact]
        public void Static_config_service_has_no_instance_adapter_yet()
        {
            string adaptersDir = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "WordTools", "Services", "Adapters");
            string adapterPath = Path.GetFullPath(Path.Combine(adaptersDir, "ConfigServiceAdapter.cs"));
            Assert.False(File.Exists(adapterPath), "IConfigService adapter is not part of this convergence pass.");
        }

        [Fact]
        public void TestAboutVersionSyncedFromVersionJson()
        {
            string versionJson = ReadProjectSource("version.json");
            var match = System.Text.RegularExpressions.Regex.Match(
                versionJson,
                "\"version\"\\s*:\\s*\"([^\"]+)\"");
            Assert.True(match.Success, "version.json must contain a version field");
            string semver = match.Groups[1].Value;

            string assemblyInfo = ReadProjectSource(Path.Combine("WordTools", "Properties", "AssemblyInfo.cs"));
            Assert.Contains(
                "[assembly: AssemblyInformationalVersion(\"" + semver + "\")]",
                assemblyInfo);

            string appVersionInfo = ReadProjectSource(Path.Combine("WordTools", "AppVersionInfo.cs"));
            Assert.Contains("AssemblyInformationalVersionAttribute", appVersionInfo);
        }

        [Fact]
        public void TestRibbonXmlExposesLoggingSettingsMenu()
        {
            string ribbonXmlPath = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "WordTools", "Ribbon.xml");
            string ribbonXml = File.ReadAllText(Path.GetFullPath(ribbonXmlPath));

            Assert.Contains("menuLoggingSettings", ribbonXml);
            Assert.Contains("chkDetailedLogging", ribbonXml);
            Assert.Contains("chkBenchmarkLogging", ribbonXml);
            Assert.Contains("imageMso=\"FileProperties\"", ribbonXml);
            Assert.DoesNotContain("splitButton id=\"splitLoggingSettings\"", ribbonXml);
            Assert.DoesNotContain("toggleButton id=\"btnToggleDetailedLogging\"", ribbonXml);
            Assert.DoesNotContain("toggleButton id=\"btnToggleBenchmarkLogging\"", ribbonXml);
        }

        [Fact]
        public void TestProgressServiceSourceAvoidsPerRefreshWordActivation()
        {
            string source = ReadProjectSource(Path.Combine("WordTools", "Services", "ProgressService.cs"));

            Assert.DoesNotContain("EnsureWordWindowActive();", source);
        }

        [Fact]
        public void TestSetupScriptDocumentsSupportedMatrix()
        {
            string source = ReadProjectSource("Setup.iss");

            Assert.Contains("ARCH_X86", source);
            Assert.Contains("ARCH_X64", source);
            Assert.Contains("WordToolbox_Setup_", source);
            Assert.Contains("\"WordToolbox_Setup_\" + MyAppVersion + \"_x86\"", source);
            Assert.Contains("\"WordToolbox_Setup_\" + MyAppVersion + \"_x64\"", source);
            Assert.Contains("仅支持 64 位 Microsoft Word", source);
            Assert.Contains("暂不支持 32 位 Word、32 位 WPS、64 位 WPS", source);
        }

        [Fact]
        public void TestRegistrationScriptsRejectUnsupportedHostsAndBitness()
        {
            string ps1 = ReadProjectSource("RegisterPlugin.ps1");
            string bat = ReadProjectSource("RegisterPlugin.bat");

            Assert.Contains("仅支持 64 位 Microsoft Word", ps1);
            Assert.Contains("暂不支持 32 位 Word、32 位 WPS、64 位 WPS", ps1);
            Assert.Contains("仅支持 64 位 Microsoft Word", bat);
            Assert.Contains("暂不支持 32 位 Word、32 位 WPS、64 位 WPS", bat);
        }

        [Fact]
        public void TestInstallationGuideDocumentsRealSupportMatrix()
        {
            string guide = ReadProjectSource("INSTALLATION.md");

            Assert.Contains("64 位 Word：支持", guide);
            Assert.Contains("32 位 Word：暂不支持", guide);
            Assert.Contains("32 位 WPS：暂不支持", guide);
            Assert.Contains("64 位 WPS：暂不支持", guide);
            Assert.Contains("x86 安装包仅用于展示当前不支持的位数提示", guide);
        }

        [Fact]
        public void TestPreferCurrentRowWithoutScanningFallback()
        {
            var currentRow = new ImageRowAvailability(40, ImageCellAvailability.Available, ImageCellAvailability.Available);

            int rowIndex = ImageRowPlanner.FindPreferredPairRow(
                currentRow,
                GetThrowingFallbackRows(),
                0);

            Assert.Equal(40, rowIndex);
        }

        [Fact]
        public void TestFallbackScanFindsNextValidRow()
        {
            var currentRow = new ImageRowAvailability(50, ImageCellAvailability.Available, ImageCellAvailability.Merged);
            var fallbackRows = new[]
            {
                new ImageRowAvailability(51, ImageCellAvailability.Available, ImageCellAvailability.Blocked),
                new ImageRowAvailability(52, ImageCellAvailability.Available, ImageCellAvailability.Available)
            };

            int rowIndex = ImageRowPlanner.FindPreferredPairRow(
                currentRow,
                fallbackRows,
                -1);

            Assert.Equal(52, rowIndex);
        }

        [Fact]
        public void TestSkipMergedRow()
        {
            var rows = new[]
            {
                new ImageRowAvailability(10, ImageCellAvailability.Available, ImageCellAvailability.Merged),
                new ImageRowAvailability(11, ImageCellAvailability.Available, ImageCellAvailability.Available)
            };

            int index = ImageRowPlanner.FindNextPairRow(rows);
            Assert.Equal(1, index);
            Assert.Equal(11, rows[index].RowIndex);
        }

        [Fact]
        public void TestSkipBlockedRow()
        {
            var rows = new[]
            {
                new ImageRowAvailability(20, ImageCellAvailability.Available, ImageCellAvailability.Blocked),
                new ImageRowAvailability(21, ImageCellAvailability.Available, ImageCellAvailability.Available)
            };

            int index = ImageRowPlanner.FindNextPairRow(rows);
            Assert.Equal(1, index);
        }

        [Fact]
        public void TestNoPairRow()
        {
            var rows = new[]
            {
                new ImageRowAvailability(30, ImageCellAvailability.Merged, ImageCellAvailability.Merged),
                new ImageRowAvailability(31, ImageCellAvailability.Available, ImageCellAvailability.Blocked)
            };

            int index = ImageRowPlanner.FindNextPairRow(rows);
            Assert.Equal(-1, index);
        }

        [Fact]
        public void TestOverwriteImageCellRemainsEligible()
        {
            var currentRow = new ImageRowAvailability(60, ImageCellAvailability.OverwriteImage, ImageCellAvailability.Available);

            int rowIndex = ImageRowPlanner.FindPreferredPairRow(
                currentRow,
                GetThrowingFallbackRows(),
                -1);

            Assert.Equal(60, rowIndex);
        }

        [Fact]
        public void TestOverwriteTextCellRemainsEligible()
        {
            var currentRow = new ImageRowAvailability(70, ImageCellAvailability.Available, ImageCellAvailability.OverwriteText);

            int rowIndex = ImageRowPlanner.FindPreferredPairRow(
                currentRow,
                GetThrowingFallbackRows(),
                -1);

            Assert.Equal(70, rowIndex);
        }

        [Fact]
        public void TestOverwriteStatesRequireWarnings()
        {
            Assert.True(ImageRowPlanner.RequiresOverwriteWarning(ImageCellAvailability.OverwriteImage),
                "existing images should produce an overwrite warning");
            Assert.True(ImageRowPlanner.RequiresOverwriteWarning(ImageCellAvailability.OverwriteText),
                "existing text should produce an overwrite warning");
            Assert.False(ImageRowPlanner.RequiresOverwriteWarning(ImageCellAvailability.Available),
                "empty reusable cells should not produce overwrite warnings");
            Assert.False(ImageRowPlanner.RequiresOverwriteWarning(ImageCellAvailability.Merged),
                "merged cells should stay in the bypass flow rather than overwrite warnings");
        }

        [Fact]
        public void TestSecondCellConflictsRetryCurrentImage()
        {
            Assert.True(ImageRowPlanner.ShouldRetryCurrentImage(ImageCellAvailability.Merged),
                "merged second cell should retry current image on a later row");
            Assert.True(ImageRowPlanner.ShouldRetryCurrentImage(ImageCellAvailability.Blocked),
                "blocked second cell should retry current image on a later row");
            Assert.False(ImageRowPlanner.ShouldRetryCurrentImage(ImageCellAvailability.Available),
                "available second cell should not trigger retry");
            Assert.False(ImageRowPlanner.ShouldRetryCurrentImage(ImageCellAvailability.OverwriteImage),
                "overwrite-eligible second cell should not trigger retry");
        }

        [Fact]
        public void TestTableServiceSourceHasNoFixedSearchWindow()
        {
            string source = ReadProjectSource(Path.Combine("WordTools", "Services", "TableService.cs"));

            Assert.DoesNotContain("startRow + 20", source);
        }

        [Fact]
        public void TestTableServiceSourceAvoidsWholeRowCellCountMergeHeuristic()
        {
            string source = ReadProjectSource(Path.Combine("WordTools", "Services", "TableService.cs"));

            Assert.DoesNotContain("mergeLeftCol = 1;", source);
        }

        [Fact]
        public void TestImageServiceSourceClearsFloatingShapes()
        {
            string source = ReadProjectSource(Path.Combine("WordTools", "Services", "ImageService.cs"));

            Assert.True(
                source.Contains("Document.Shapes") || source.Contains(".Shapes"),
                "image overwrite path should consider floating shapes, not only inline shapes");
        }

        [Fact]
        public void TestImageServiceBatchPathUsesContextAwareOverwriteClear()
        {
            string source = ReadProjectSource(Path.Combine("WordTools", "Services", "ImageService.cs"));

            Assert.Contains("ClearCellContentForOverwrite(targetCell, context);", source);
        }

        private static IEnumerable<ImageRowAvailability> GetThrowingFallbackRows()
        {
            return new ThrowingEnumerable();
        }

        private static string ReadProjectSource(string relativePath)
        {
            string sourcePath = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", relativePath);
            return File.ReadAllText(Path.GetFullPath(sourcePath));
        }

        private static IEnumerable<Control> GetAllControls(Control root)
        {
            foreach (Control child in root.Controls)
            {
                yield return child;

                foreach (Control descendant in GetAllControls(child))
                {
                    yield return descendant;
                }
            }
        }

        [Fact]
        public void ProgressServiceSourceDoesNotReferenceGlobalsOrMessageBox()
        {
            string source = ReadProjectSource(Path.Combine("WordTools", "Services", "ProgressService.cs"));
            Assert.DoesNotContain("Globals.ThisAddIn", source);
            Assert.DoesNotContain("MessageBox.Show", source);
            Assert.DoesNotContain("new ProgressForm", source);
            Assert.DoesNotContain("new FailureDetailsForm", source);
        }

        [Fact]
        public void AbstractionsExist()
        {
            Assert.True(typeof(WordTools.Services.Abstractions.IProgressReporter).IsInterface);
            Assert.True(typeof(WordTools.Services.Abstractions.IFailureDetailsPresenter).IsInterface);
            Assert.True(typeof(WordTools.Services.Abstractions.INotificationService).IsInterface);
            Assert.True(typeof(WordTools.Services.Abstractions.IDocumentContext).IsInterface);
            Assert.True(typeof(WordTools.Services.Abstractions.IWordApplicationContext).IsInterface);
        }

        private sealed class ThrowingEnumerable : IEnumerable<ImageRowAvailability>
        {
            public IEnumerator<ImageRowAvailability> GetEnumerator()
            {
                throw new InvalidOperationException("fallback rows should not be enumerated for the fast path");
            }

            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }
        }
    }
}
