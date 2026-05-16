using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using WordTools.Forms;
using WordTools.Services;

static class Program
{
    private static int _failed;

    private static void Main()
    {
        Run("Progress button switches to close after completion", TestProgressButtonSwitchesToCloseAfterCompletion);
        Run("Benchmark log writes header once", TestBenchmarkLogWritesHeaderOnce);
        Run("Benchmark log escapes CSV fields", TestBenchmarkLogEscapesCsvFields);
        Run("Logging options disable benchmark without detailed log", TestLoggingOptionsDisableBenchmarkWithoutDetailedLog);
        Run("Logging options allow benchmark with detailed log", TestLoggingOptionsAllowBenchmarkWithDetailedLog);
        Run("Insertion diagnostics report the new hot-path breakdown", TestInsertionDiagnosticsReportHotPathBreakdown);
        Run("Insertion diagnostics report validation, sizing, and progress UI time", TestInsertionDiagnosticsReportExtendedHotPathBreakdown);
        Run("Floating shape index finds anchors within the current cell range", TestFloatingShapeIndexFindsAnchorsWithinCellRange);
        Run("Batch context reuses floating shape index until invalidated", TestBatchContextReusesFloatingShapeIndexUntilInvalidated);
        Run("Batch context caches row availability until cleared", TestBatchContextCachesRowAvailabilityUntilCleared);
        Run("Existing image cells remain eligible for overwrite", TestOverwriteImageCellRemainsEligible);
        Run("Existing text cells remain eligible for overwrite", TestOverwriteTextCellRemainsEligible);
        Run("Overwrite states require warning prompts", TestOverwriteStatesRequireWarnings);
        Run("Second-cell conflicts retry the current image", TestSecondCellConflictsRetryCurrentImage);
        Run("Table service no longer uses a fixed 20-row search window", TestTableServiceSourceHasNoFixedSearchWindow);
        Run("Table service no longer marks entire merged rows unavailable by cell-count alone", TestTableServiceSourceAvoidsWholeRowCellCountMergeHeuristic);
        Run("Image service clears floating shapes during overwrite", TestImageServiceSourceClearsFloatingShapes);
        Run("Image service fast path uses context-aware overwrite clear", TestImageServiceBatchPathUsesContextAwareOverwriteClear);
        Run("Summary is required when only warnings exist", TestSummaryRequiredForWarningsWithoutFailures);
        Run("Summary message includes merged and overwrite sections", TestSummaryMessageIncludesMergedAndOverwriteSections);
        Run("Summary message explains Yes/No details action", TestSummaryMessageExplainsDetailsAction);
        Run("Failure details buttons use explicit labels", TestFailureDetailsButtonsUseExplicitLabels);
        Run("Progress service source contains no mojibake markers", TestProgressServiceSourceHasNoMojibake);
        Run("Insert photos form no longer contains logging controls", TestInsertPhotosFormSourceHasNoLoggingControls);
        Run("Insert photos form closes before launching insertion work", TestInsertPhotosFormClosesBeforeLaunchingInsertionWork);
        Run("Ribbon XML exposes a logging settings menu", TestRibbonXmlExposesLoggingSettingsMenu);
        Run("Progress service no longer forces Word activation on every progress refresh", TestProgressServiceSourceAvoidsPerRefreshWordActivation);
        Run("Inno Setup exposes only 64-bit Word as supported", TestSetupScriptDocumentsSupportedMatrix);
        Run("Registration scripts clearly reject unsupported hosts and bitness", TestRegistrationScriptsRejectUnsupportedHostsAndBitness);
        Run("Installation guide documents the real support matrix", TestInstallationGuideDocumentsRealSupportMatrix);
        Run("Prefer current row without scanning fallback", TestPreferCurrentRowWithoutScanningFallback);
        Run("Fallback scan finds next valid row", TestFallbackScanFindsNextValidRow);
        Run("Skip row when second cell is merged", TestSkipMergedRow);
        Run("Skip row when second cell is blocked", TestSkipBlockedRow);
        Run("Return -1 when no pair row exists", TestNoPairRow);

        if (_failed > 0)
        {
            Environment.Exit(1);
        }
    }

    private static void Run(string name, Action test)
    {
        try
        {
            test();
            Console.WriteLine("[PASS] " + name);
        }
        catch (Exception ex)
        {
            _failed++;
            Console.WriteLine("[FAIL] " + name + ": " + ex.Message);
        }
    }

    private static void TestSkipMergedRow()
    {
        var rows = new[]
        {
            new ImageRowAvailability(10, ImageCellAvailability.Available, ImageCellAvailability.Merged),
            new ImageRowAvailability(11, ImageCellAvailability.Available, ImageCellAvailability.Available)
        };

        int index = ImageRowPlanner.FindNextPairRow(rows);
        AssertEqual(1, index, "planner should skip rows that cannot host two images");
        AssertEqual(11, rows[index].RowIndex, "planner should move to the next fully available row");
    }

    private static void TestSkipBlockedRow()
    {
        var rows = new[]
        {
            new ImageRowAvailability(20, ImageCellAvailability.Available, ImageCellAvailability.Blocked),
            new ImageRowAvailability(21, ImageCellAvailability.Available, ImageCellAvailability.Available)
        };

        int index = ImageRowPlanner.FindNextPairRow(rows);
        AssertEqual(1, index, "planner should not start a row that cannot finish two-image layout");
    }

    private static void TestNoPairRow()
    {
        var rows = new[]
        {
            new ImageRowAvailability(30, ImageCellAvailability.Merged, ImageCellAvailability.Merged),
            new ImageRowAvailability(31, ImageCellAvailability.Available, ImageCellAvailability.Blocked)
        };

        int index = ImageRowPlanner.FindNextPairRow(rows);
        AssertEqual(-1, index, "planner should report that no valid two-image row exists");
    }

    private static void TestPreferCurrentRowWithoutScanningFallback()
    {
        var currentRow = new ImageRowAvailability(40, ImageCellAvailability.Available, ImageCellAvailability.Available);

        int rowIndex = ImageRowPlanner.FindPreferredPairRow(
            currentRow,
            GetThrowingFallbackRows(),
            0);

        AssertEqual(40, rowIndex, "planner should return current row immediately when it already supports a full pair");
    }

    private static void TestFallbackScanFindsNextValidRow()
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

        AssertEqual(52, rowIndex, "planner should scan fallback rows only when current row cannot host two images");
    }

    private static void TestOverwriteImageCellRemainsEligible()
    {
        var currentRow = new ImageRowAvailability(60, ImageCellAvailability.OverwriteImage, ImageCellAvailability.Available);

        int rowIndex = ImageRowPlanner.FindPreferredPairRow(
            currentRow,
            GetThrowingFallbackRows(),
            -1);

        AssertEqual(60, rowIndex, "rows with an existing image should remain eligible when overwrite is allowed");
    }

    private static void TestOverwriteTextCellRemainsEligible()
    {
        var currentRow = new ImageRowAvailability(70, ImageCellAvailability.OverwriteText, ImageCellAvailability.OverwriteText);

        int rowIndex = ImageRowPlanner.FindPreferredPairRow(
            currentRow,
            GetThrowingFallbackRows(),
            -1);

        AssertEqual(70, rowIndex, "rows with existing text should remain eligible when overwrite is allowed");
    }

    private static void TestOverwriteStatesRequireWarnings()
    {
        AssertTrue(ImageRowPlanner.RequiresOverwriteWarning(ImageCellAvailability.OverwriteImage),
            "existing images should produce an overwrite warning");
        AssertTrue(ImageRowPlanner.RequiresOverwriteWarning(ImageCellAvailability.OverwriteText),
            "existing text should produce an overwrite warning");
        AssertTrue(!ImageRowPlanner.RequiresOverwriteWarning(ImageCellAvailability.Available),
            "empty reusable cells should not produce overwrite warnings");
        AssertTrue(!ImageRowPlanner.RequiresOverwriteWarning(ImageCellAvailability.Merged),
            "merged cells should stay in the bypass flow rather than overwrite warnings");
    }

    private static void TestSecondCellConflictsRetryCurrentImage()
    {
        AssertTrue(ImageRowPlanner.ShouldRetryCurrentImage(ImageCellAvailability.Merged),
            "merged second cell should retry current image on a later row");
        AssertTrue(ImageRowPlanner.ShouldRetryCurrentImage(ImageCellAvailability.Blocked),
            "blocked second cell should retry current image on a later row");
        AssertTrue(!ImageRowPlanner.ShouldRetryCurrentImage(ImageCellAvailability.Available),
            "available second cell should not trigger retry");
        AssertTrue(!ImageRowPlanner.ShouldRetryCurrentImage(ImageCellAvailability.OverwriteImage),
            "overwrite-eligible second cell should not trigger retry");
    }

    private static void TestTableServiceSourceHasNoFixedSearchWindow()
    {
        string source = ReadProjectSource(Path.Combine("WordTools", "Services", "TableService.cs"));

        AssertTrue(source.IndexOf("startRow + 20", StringComparison.Ordinal) < 0,
            "table row search should not stop after an arbitrary 20-row window");
    }

    private static void TestTableServiceSourceAvoidsWholeRowCellCountMergeHeuristic()
    {
        string source = ReadProjectSource(Path.Combine("WordTools", "Services", "TableService.cs"));

        AssertTrue(source.IndexOf("mergeLeftCol = 1;", StringComparison.Ordinal) < 0,
            "table service should not mark an entire row merged by forcing the merge origin to column 1");
    }

    private static void TestImageServiceSourceClearsFloatingShapes()
    {
        string source = ReadProjectSource(Path.Combine("WordTools", "Services", "ImageService.cs"));

        AssertTrue(
            source.IndexOf("Document.Shapes", StringComparison.Ordinal) >= 0
            || source.IndexOf(".Shapes", StringComparison.Ordinal) >= 0,
            "image overwrite path should consider floating shapes, not only inline shapes");
    }

    private static void TestImageServiceBatchPathUsesContextAwareOverwriteClear()
    {
        string source = ReadProjectSource(Path.Combine("WordTools", "Services", "ImageService.cs"));

        AssertTrue(source.IndexOf("ClearCellContentForOverwrite(targetCell, context);", StringComparison.Ordinal) >= 0,
            "batch insertion path should reuse the context-aware overwrite clear helper");
    }

    private static void TestLoggingOptionsDisableBenchmarkWithoutDetailedLog()
    {
        LoggingOptionsState state = LoggingOptionsStateController.Normalize(false, true);

        AssertTrue(!state.DetailedLoggingEnabled, "detailed logging should stay disabled");
        AssertTrue(!state.BenchmarkLoggingEnabled, "benchmark logging should be cleared when detailed logging is disabled");
        AssertTrue(!LoggingOptionsStateController.ShouldWriteBenchmarkLog(false, true),
            "benchmark log should not write when detailed logging is disabled");
        AssertTrue(!LoggingOptionsStateController.ShouldShowDetailedLog(false),
            "diagnostic detail should stay hidden by default");
    }

    private static void TestLoggingOptionsAllowBenchmarkWithDetailedLog()
    {
        LoggingOptionsState state = LoggingOptionsStateController.Normalize(true, true);

        AssertTrue(state.DetailedLoggingEnabled, "detailed logging should remain enabled");
        AssertTrue(state.BenchmarkLoggingEnabled, "benchmark logging should remain enabled when detailed logging is enabled");
        AssertTrue(LoggingOptionsStateController.ShouldWriteBenchmarkLog(true, true),
            "benchmark log should write only when both toggles are enabled");
        AssertTrue(LoggingOptionsStateController.ShouldShowDetailedLog(true),
            "diagnostic detail should appear when detailed logging is enabled");
    }

    private static void TestInsertionDiagnosticsReportHotPathBreakdown()
    {
        var diagnostics = new InsertionPerformanceDiagnostics();
        diagnostics.RecordCellAvailabilityCheck(12);
        diagnostics.RecordCellAvailabilityCheck(8);
        diagnostics.RecordFloatingShapeLookup(5);
        diagnostics.RecordOverwriteClear(7);
        diagnostics.RecordAddPicture(42);
        diagnostics.RecordDescriptionWrite(11);

        string detail = diagnostics.BuildDetailedLog();

        AssertTrue(detail.Contains("行可用性检查: 20ms (2次)"),
            "diagnostics should report aggregated row-availability checks");
        AssertTrue(detail.Contains("浮动图片探测: 5ms (1次)"),
            "diagnostics should report floating-shape lookup time");
        AssertTrue(detail.Contains("覆盖清理: 7ms (1次)"),
            "diagnostics should report overwrite-clear time");
        AssertTrue(detail.Contains("AddPicture: 42ms (1次)"),
            "diagnostics should report AddPicture time separately");
        AssertTrue(detail.Contains("描述/编号写入: 11ms (1次)"),
            "diagnostics should report description and numbering writes");
    }

    private static void TestInsertionDiagnosticsReportExtendedHotPathBreakdown()
    {
        var diagnostics = new InsertionPerformanceDiagnostics();
        diagnostics.RecordCellValidation(9);
        diagnostics.RecordPictureSizing(14);
        diagnostics.RecordProgressUi(6);

        string detail = diagnostics.BuildDetailedLog();

        AssertTrue(detail.Contains("Cell validation: 9ms (1"),
            "diagnostics should report cell validation time separately");
        AssertTrue(detail.Contains("Picture sizing: 14ms (1"),
            "diagnostics should report picture sizing time separately");
        AssertTrue(detail.Contains("Progress UI: 6ms (1"),
            "diagnostics should report progress UI overhead separately");
    }

    private static void TestFloatingShapeIndexFindsAnchorsWithinCellRange()
    {
        FloatingShapeIndex index = FloatingShapeIndex.Create(new[]
        {
            new FloatingShapeAnchor(10, 18),
            new FloatingShapeAnchor(30, 38),
            new FloatingShapeAnchor(60, 70)
        });

        AssertTrue(index.HasShapeInRange(1, 20),
            "index should find anchors fully contained in the current cell range");
        AssertTrue(index.HasShapeInRange(25, 40),
            "index should find later anchors in a different cell range");
        AssertTrue(!index.HasShapeInRange(19, 29),
            "index should not report anchors outside the current cell range");
        AssertTrue(!index.HasShapeInRange(61, 65),
            "index should require the whole anchor to fit in the cell range");
    }

    private static void TestBatchContextReusesFloatingShapeIndexUntilInvalidated()
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

        AssertEqual(1, factoryCalls, "floating shape index should be created only once while cache is valid");
        AssertTrue(object.ReferenceEquals(first, second), "context should reuse the cached floating shape index");

        context.InvalidateFloatingShapeIndex();

        FloatingShapeIndex third = context.GetOrCreateFloatingShapeIndex(() =>
        {
            factoryCalls++;
            return FloatingShapeIndex.Create(new[] { new FloatingShapeAnchor(20, 29) });
        });

        AssertEqual(2, factoryCalls, "invalidating the cache should rebuild the floating shape index");
        AssertTrue(!object.ReferenceEquals(first, third), "rebuilt floating shape index should replace the old cached instance");
    }

    private static void TestBatchContextCachesRowAvailabilityUntilCleared()
    {
        var context = new ImageInsertionBatchContext();
        var row = new ImageRowAvailability(88, ImageCellAvailability.Available, ImageCellAvailability.OverwriteText);

        AssertTrue(!context.TryGetCachedRowAvailability(88, out _),
            "empty context should not report a cached row");

        context.CacheRowAvailability(row);

        AssertTrue(context.TryGetCachedRowAvailability(88, out var cachedRow),
            "cached row should be returned for later lookups");
        AssertEqual(ImageCellAvailability.OverwriteText, cachedRow.RightCell,
            "cached row should preserve the original right-cell availability");

        context.ClearRowAvailability();

        AssertTrue(!context.TryGetCachedRowAvailability(88, out _),
            "clearing the cache should remove stored row availability");
    }

    private static void TestSummaryRequiredForWarningsWithoutFailures()
    {
        AssertTrue(InsertionSummaryFormatter.ShouldShowSummary(
                0,
                new List<int> { 3 },
                new List<string>()),
            "merged-cell warnings alone should still trigger the unified summary");

        AssertTrue(InsertionSummaryFormatter.ShouldShowSummary(
                0,
                new List<int>(),
                new List<string> { "第7行第1列已有图片，已覆盖插入新图片" }),
            "overwrite warnings alone should still trigger the unified summary");
    }

    private static void TestSummaryMessageIncludesMergedAndOverwriteSections()
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

        AssertTrue(message.Contains("失败: 1 张"), "summary should include failure count");
        AssertTrue(message.Contains("合并单元格绕开: 1 处"), "summary should include merged-cell count");
        AssertTrue(message.Contains("覆盖插图提示: 1 处"), "summary should include overwrite count");
        AssertTrue(message.Contains("第8行"), "summary should preview merged-cell locations");
        AssertTrue(message.Contains("第7行第1列已有文本"), "summary should preview overwrite warnings");
    }

    private static void TestSummaryMessageExplainsDetailsAction()
    {
        string message = InsertionSummaryFormatter.BuildDetailsPrompt();

        AssertTrue(message.Contains("是(Y)"), "details prompt should explain the meaning of the Yes button");
        AssertTrue(message.Contains("查看详情"), "details prompt should explicitly mention viewing details");
        AssertTrue(message.Contains("否(N)"), "details prompt should explain the meaning of the No button");
    }

    private static void TestFailureDetailsButtonsUseExplicitLabels()
    {
        using var form = new FailureDetailsForm(
            new List<(string fileName, string errorReason)> { ("a.jpg", "文件不存在或已被移动") },
            new List<int> { 11 },
            new List<string> { "第7行第1列已有文本，已覆盖插入新图片" });

        var buttons = GetAllControls(form).OfType<Button>().ToList();
        var buttonTexts = buttons.Select(button => button.Text).ToList();

        AssertTrue(buttonTexts.Contains("复制详情"), "details dialog should expose a clear copy-details label");
        AssertTrue(buttonTexts.Contains("关闭窗口"), "details dialog should expose a clear close-window label");

        Button copyButton = buttons.Single(button => button.Text == "复制详情");
        Button closeButton = buttons.Single(button => button.Text == "关闭窗口");
        AssertTrue(copyButton.Width >= 96, "copy-details button should be wide enough to avoid clipping");
        AssertTrue(closeButton.Width >= 96, "close-window button should be wide enough to avoid clipping");
        AssertTrue(copyButton.Height >= 32, "copy-details button should be tall enough to avoid clipping");
        AssertTrue(closeButton.Height >= 32, "close-window button should be tall enough to avoid clipping");
    }

    private static void TestProgressServiceSourceHasNoMojibake()
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
            AssertTrue(source.IndexOf(marker, StringComparison.Ordinal) < 0,
                "progress service still contains mojibake marker: " + marker);
        }
    }

    private static void TestInsertPhotosFormSourceHasNoLoggingControls()
    {
        string sourcePath = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "WordTools", "Forms", "InsertPhotosForm.cs");
        string source = File.ReadAllText(Path.GetFullPath(sourcePath));

        AssertTrue(source.IndexOf("chkDetailedLogging", StringComparison.Ordinal) < 0,
            "insert photos form should no longer declare detailed logging controls");
        AssertTrue(source.IndexOf("chkBenchmarkLogging", StringComparison.Ordinal) < 0,
            "insert photos form should no longer declare benchmark logging controls");
        AssertTrue(source.IndexOf("CreateLoggingOptionsSection", StringComparison.Ordinal) < 0,
            "insert photos form should no longer build a logging section");
    }

    private static void TestInsertPhotosFormClosesBeforeLaunchingInsertionWork()
    {
        string sourcePath = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "WordTools", "Forms", "InsertPhotosForm.cs");
        string source = File.ReadAllText(Path.GetFullPath(sourcePath));
        string addInPath = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "WordTools", "ThisAddIn.cs");
        string addInSource = File.ReadAllText(Path.GetFullPath(addInPath));

        AssertTrue(source.IndexOf("Hide();", StringComparison.Ordinal) < 0,
            "modal insert dialog should not hide itself before launching background work");
        AssertTrue(source.Contains("PendingRequest = new InsertPhotosRequest"),
            "insert dialog should hand off a request object instead of running insertion directly");
        AssertTrue(source.Contains("DialogResult = DialogResult.OK;"),
            "insert dialog should close with an OK result after preparing the request");
        AssertTrue(addInSource.Contains("ExecuteInsertPhotosRequestDeferred(pendingRequest);"),
            "add-in should defer insertion until after the modal dialog has fully closed");
        AssertTrue(addInSource.Contains("BeginInvoke(new Action(() =>"),
            "add-in should schedule insertion on the next UI message loop turn");
    }

    private static void TestRibbonXmlExposesLoggingSettingsMenu()
    {
        string ribbonXmlPath = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "WordTools", "Ribbon.xml");
        string ribbonXml = File.ReadAllText(Path.GetFullPath(ribbonXmlPath));

        AssertTrue(ribbonXml.Contains("menuLoggingSettings"),
            "ribbon should expose a top-level logging settings menu");
        AssertTrue(ribbonXml.Contains("chkDetailedLogging"),
            "logging settings menu should contain the detailed logging checkbox");
        AssertTrue(ribbonXml.Contains("chkBenchmarkLogging"),
            "logging settings menu should contain the benchmark logging checkbox");
        AssertTrue(ribbonXml.Contains("imageMso=\"FileProperties\""),
            "logging settings entry should use a visible configuration-style icon");
        AssertTrue(ribbonXml.IndexOf("splitButton id=\"splitLoggingSettings\"", StringComparison.Ordinal) < 0,
            "logging settings should avoid the split button structure that can break Ribbon loading");
        AssertTrue(ribbonXml.IndexOf("toggleButton id=\"btnToggleDetailedLogging\"", StringComparison.Ordinal) < 0,
            "detailed logging should no longer be rendered as a top-level toggle button");
        AssertTrue(ribbonXml.IndexOf("toggleButton id=\"btnToggleBenchmarkLogging\"", StringComparison.Ordinal) < 0,
            "benchmark logging should no longer be rendered as a top-level toggle button");
    }

    private static void TestProgressServiceSourceAvoidsPerRefreshWordActivation()
    {
        string source = ReadProjectSource(Path.Combine("WordTools", "Services", "ProgressService.cs"));

        AssertTrue(source.IndexOf("EnsureWordWindowActive();", StringComparison.Ordinal) < 0,
            "progress refresh should not force Word activation on every UI update");
    }

    private static void TestSetupScriptDocumentsSupportedMatrix()
    {
        string source = ReadProjectSource("Setup.iss");

        AssertTrue(source.Contains("ARCH_X86") && source.Contains("ARCH_X64"),
            "setup script should define separate x86 and x64 build switches");
        AssertTrue(source.Contains("WordToolbox_Setup_x86") && source.Contains("WordToolbox_Setup_x64"),
            "setup script should produce architecture-specific output package names");
        AssertTrue(source.Contains("仅支持 64 位 Microsoft Word"),
            "setup script should state that only 64-bit Word is supported");
        AssertTrue(source.Contains("暂不支持 32 位 Word、32 位 WPS、64 位 WPS"),
            "setup script should explicitly warn about unsupported environments");
    }

    private static void TestRegistrationScriptsRejectUnsupportedHostsAndBitness()
    {
        string ps1 = ReadProjectSource("RegisterPlugin.ps1");
        string bat = ReadProjectSource("RegisterPlugin.bat");

        AssertTrue(ps1.Contains("仅支持 64 位 Microsoft Word"),
            "PowerShell registration script should explain the only supported host");
        AssertTrue(ps1.Contains("暂不支持 32 位 Word、32 位 WPS、64 位 WPS"),
            "PowerShell registration script should reject unsupported hosts and bitness");
        AssertTrue(bat.Contains("仅支持 64 位 Microsoft Word"),
            "batch registration script should explain the only supported host");
        AssertTrue(bat.Contains("暂不支持 32 位 Word、32 位 WPS、64 位 WPS"),
            "batch registration script should reject unsupported hosts and bitness");
    }

    private static void TestInstallationGuideDocumentsRealSupportMatrix()
    {
        string guide = ReadProjectSource("INSTALLATION.md");

        AssertTrue(guide.Contains("64 位 Word：支持"),
            "installation guide should mark 64-bit Word as supported");
        AssertTrue(guide.Contains("32 位 Word：暂不支持"),
            "installation guide should mark 32-bit Word as unsupported");
        AssertTrue(guide.Contains("32 位 WPS：暂不支持") && guide.Contains("64 位 WPS：暂不支持"),
            "installation guide should mark WPS variants as unsupported");
        AssertTrue(guide.Contains("x86 安装包仅用于展示当前不支持的位数提示"),
            "installation guide should clarify that the x86 package is not a supported delivery target");
    }

    private static void TestBenchmarkLogWritesHeaderOnce()
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
            AssertEqual(3, lines.Length, "benchmark log should contain one header plus two rows");
            AssertTrue(lines[0].Contains("timestamp_utc"), "header should be written when file is first created");
        }
        finally
        {
            if (File.Exists(tempFile))
            {
                File.Delete(tempFile);
            }
        }
    }

    private static void TestProgressButtonSwitchesToCloseAfterCompletion()
    {
        var controller = new ProgressFormStateController();

        ProgressButtonAction firstClick = controller.HandleButtonClick();
        AssertEqual(ProgressButtonAction.CancelRequested, firstClick, "first click should request cancellation during active progress");
        AssertTrue(controller.IsCancelled, "controller should record cancellation request");

        controller.MarkCompleted();

        ProgressButtonAction completionClick = controller.HandleButtonClick();
        AssertEqual(ProgressButtonAction.CloseRequested, completionClick, "button should become a close action after completion");
        AssertEqual("关闭", controller.ButtonText, "completed state should expose close text");
        AssertTrue(controller.IsButtonEnabled, "completed state should re-enable the primary button");
    }

    private static void TestBenchmarkLogEscapesCsvFields()
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
            AssertTrue(lines[1].Contains("\"D:\\Data,Set\\Input\""), "commas should be quoted");
            AssertTrue(lines[1].Contains("\"line1 line2 \"\"quoted\"\"\""), "newlines should be flattened and quotes doubled");
        }
        finally
        {
            if (File.Exists(tempFile))
            {
                File.Delete(tempFile);
            }
        }
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

    private static void AssertEqual<T>(T expected, T actual, string message)
    {
        if (!EqualityComparer<T>.Default.Equals(expected, actual))
        {
            throw new InvalidOperationException(message + $" (expected: {expected}, actual: {actual})");
        }
    }

    private static void AssertTrue(bool condition, string message)
    {
        if (!condition)
        {
            throw new InvalidOperationException(message);
        }
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
