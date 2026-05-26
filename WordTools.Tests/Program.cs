using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
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
        Run("Inno Setup keeps the single public installer while limiting live registration to verified Word x64", TestSetupScriptDocumentsSinglePublicInstallerAndVerifiedLiveRegistrationFlow);
        Run("Registration scripts reject unsupported WPS and x86 targets through the shared-core wrapper", TestRegistrationScriptsRejectUnsupportedHostsAndBitnessThroughWrapper);
        Run("Registration wrapper stops hard-blocking Word x86 before shared-core support evaluation", TestRegistrationWrapperStopsHardBlockingWordX86BeforeSharedCoreEvaluation);
        Run("Installation guide documents the rolled-back real support matrix", TestInstallationGuideDocumentsSingleInstallerRealSupportMatrix);
        Run("Installation guide removes the old x86 placeholder package narrative", TestInstallationGuideRemovesOldX86PlaceholderPackageNarrative);
        Run("README points to the unified installer flow and Word-only support boundary", TestReadmePointsToUnifiedInstallerFlow);
        Run("Shared core records installer state and reuses it for live uninstall fallback", TestSharedCoreRecordsInstallerStateAndUsesItForLiveUninstallFallback);
        Run("Shared core prepares Word x86 installer-state rehydration and removes Word x64-only live wording", TestSharedCorePreparesWordX86InstallerStateAndMessages);
        Run("Shared core rehydrates installer-state hosts through the support matrix", TestSharedCoreRehydratesInstallerStateHostsThroughSupportMatrix);
        Run("Shared core parses installer-state host labels generically", TestSharedCoreParsesInstallerStateHostLabelsGenerically);
        Run("Shared core rehydrates installer-state execution metadata for diagnostics", TestSharedCoreRehydratesInstallerStateExecutionMetadataForDiagnostics);
        Run("Shared core preserves version lines through installer-state rehydration", TestSharedCorePreservesVersionLinesThroughInstallerStateRehydration);
        Run("Manual verification checklist covers installer state and WPS residual cleanup", TestManualVerificationChecklistCoversInstallerStateAndWpsResidualCleanup);
        Run("Manual verification checklist records the WPS UI failure and keeps WPS out of the supported boundary", TestManualVerificationChecklistRecordsWpsUiFailureAndKeepsWpsUnsupported);
        Run("WPS recon evidence documents native package metadata and keeps AddinsWl unvalidated", TestWpsReconEvidenceDocumentsNativePackageMetadata);
        Run("WPS live handlers remain implemented behind a disabled feature gate", TestWpsLiveHandlersAreImplementedBehindDisabledFeatureGate);
        Run("Probe core scaffold exists without replacing current registration path", TestProbeCoreScaffoldExistsWithoutChangingRegistrationPath);
        Run("Support matrix keeps Word x64 supported while rolling WPS x86 back to planned", TestSupportMatrixPromotesWordX64AndKeepsOthersPlanned);
        Run("Support matrix declares validation stages and activation routes", TestSupportMatrixDeclaresValidationStagesAndActivationRoutes);
        Run("Invoke-WpsAddinsWlExperiment is defined with mandatory parameters and Restore switch", TestInvokeWpsAddinsWlExperimentDefined);
        Run("Invoke-WpsAddinsWlExperiment backup flow verifies .reg existence and tracks entry counts", TestInvokeWpsAddinsWlExperimentBackupFlow);
        Run("Support matrix has ui-experiment state for WPS x86", TestSupportMatrixHasUiExperimentState);
        Run("Register entrypoint reroutes through shared core while preserving current Word-only support guardrails", TestReroutedRegistrationEntryPointPreservesWordOnlySupportGuardrails);
        Run("Unified core exposes register and unregister preview modes alongside the rerouted live entrypoint", TestUnifiedCoreExposesRegisterAndUnregisterPreviewModesAlongsideReroutedLiveEntrypoint);
        Run("Live register mode fails fast with explicit administrator guidance when not elevated", TestLiveRegisterModeFailsFastWithExplicitAdministratorGuidance);
        Run("Live register mode writes failure payloads and summaries when live execution fails", TestLiveRegisterModeWritesFailurePayloadAndSummaryWhenLiveExecutionFails);
        Run("Shared core source supports installer-driven self elevation", TestSharedCoreSourceSupportsInstallerDrivenSelfElevation);
        Run("Shared core source captures external tool output before returning live objects", TestSharedCoreSourceCapturesExternalToolOutputBeforeReturningLiveObjects);
        Run("Probe output structure includes support reason and registration view", TestProbeOutputStructureIncludesSupportReasonAndRegistrationView);
        Run("Probe output structure includes validation stage and activation route", TestProbeOutputStructureIncludesValidationStageAndActivationRoute);
        Run("Probe output structure includes host state model fields", TestProbeOutputStructureIncludesHostStateModelFields);
        Run("Probe output structure includes WPS reconnaissance details when WPS is detected", TestProbeOutputStructureIncludesWpsReconnaissanceDetails);
        Run("Probe output structure includes top-level support summary", TestProbeOutputStructureIncludesTopLevelSupportSummary);
        Run("Probe output structure includes missing and ambiguous host summaries", TestProbeOutputStructureIncludesMissingAndAmbiguousHostSummaries);
        Run("Probe output structure includes ambiguity reason details", TestProbeOutputStructureIncludesAmbiguityReasonDetails);
        Run("Probe supports saving JSON output to a file", TestProbeSupportsSavingJsonOutputToFile);
        Run("Probe supports writing summary text with host-state details", TestProbeSupportsWritingSummaryTextWithHostStateDetails);
        Run("Probe supports attaching an evidence label", TestProbeSupportsAttachingEvidenceLabel);
        Run("Probe supports appending evidence markdown rows", TestProbeSupportsAppendingEvidenceMarkdownRows);
        Run("Evidence markdown scaffold uses validation-stage structure", TestEvidenceMarkdownScaffoldUsesValidationStageStructure);
        Run("Evidence markdown scaffold seeds the current support matrix baseline", TestEvidenceMarkdownScaffoldSeedsCurrentSupportMatrixBaseline);
        Run("Host detection matrix uses pending-matrix and evidence-log sections", TestHostDetectionMatrixUsesPendingAndEvidenceSections);
        Run("Plan mode exposes a dry-run registration plan for supported Word x64", TestPlanModeExposesDryRunRegistrationPlanForSupportedWordX64);
        Run("Plan mode respects explicit Word x86 requests when no Word x86 host is detected", TestPlanModeRespectsExplicitWordX86RequestWhenNoWordX86HostIsDetected);
        Run("Plan mode summary text exposes host-state details for diagnostics", TestPlanModeSummaryTextExposesHostStateDetailsForDiagnostics);
        Run("Unregister preview respects explicit Word x86 requests when no Word x86 host is detected", TestUnregisterPreviewRespectsExplicitWordX86RequestWhenNoWordX86HostIsDetected);
        Run("Dry-run summary text derives registrable targets from dry-run eligibility", TestDryRunSummaryTextDerivesRegistrableTargetsFromDryRunEligibility);
        Run("Unregister preview summary text exposes host-state details for diagnostics", TestUnregisterPreviewSummaryTextExposesHostStateDetailsForDiagnostics);
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

    private static void TestSetupScriptDocumentsSinglePublicInstallerAndSupportedMatrix()
    {
        string source = ReadProjectSource("Setup.iss");
        string buildScript = ReadProjectSource("build-installer.ps1");

        AssertTrue(source.Contains("OutputBaseFilename=WordToolbox_Setup", StringComparison.Ordinal),
            "setup script should emit a single public installer filename");
        AssertTrue(source.IndexOf("WordToolbox_Setup_x86", StringComparison.Ordinal) < 0,
            "setup script should no longer emit a public x86 package");
        AssertTrue(source.IndexOf("WordToolbox_Setup_x64", StringComparison.Ordinal) < 0,
            "setup script should no longer emit a public x64-only filename");
        AssertTrue(source.Contains("ArchitecturesAllowed=x86compatible", StringComparison.Ordinal),
            "setup script should keep the unified public installer runnable on both 32-bit and 64-bit Windows while Phase 1 expands toward Word x86");
        AssertTrue(source.Contains("ArchitecturesInstallIn64BitMode=x64compatible", StringComparison.Ordinal),
            "setup script should still enter 64-bit install mode when the machine supports it");
        AssertTrue(source.Contains("DefaultDirName={autopf}\\WordToolbox", StringComparison.Ordinal),
            "setup script should resolve the install root through the current platform Program Files view instead of hard-locking to 64-bit Program Files");
        AssertTrue(source.Contains("Source: \"Installer.Core.ps1\"", StringComparison.Ordinal)
            && source.Contains("Source: \"Installer.SupportMatrix.json\"", StringComparison.Ordinal),
            "setup script should bundle the shared installer core and support matrix");
        AssertTrue(source.Contains("Installer.Core.ps1\"\" -Mode Register", StringComparison.Ordinal)
            || source.Contains("Installer.Core.ps1\" -Mode Register", StringComparison.Ordinal),
            "setup script should invoke the shared installer core for registration");
        AssertTrue(source.Contains("64 位 Microsoft Word", StringComparison.Ordinal),
            "setup script should state that only 64-bit Word is supported");
        AssertTrue(source.Contains("32 位 Word", StringComparison.Ordinal)
            && source.Contains("32 位 WPS", StringComparison.Ordinal)
            && source.Contains("64 位 WPS", StringComparison.Ordinal),
            "setup script should explicitly warn about unsupported environments");
        AssertTrue(source.Contains("DetectedHostDetails:", StringComparison.Ordinal),
            "setup script should require the probe summary to include host-state details before continuing");
        AssertTrue(buildScript.Contains("WordToolbox_Setup.exe", StringComparison.Ordinal),
            "installer build helper should report a single public installer output");
        AssertTrue(buildScript.IndexOf("ARCH_X86", StringComparison.Ordinal) < 0
            && buildScript.IndexOf("ARCH_X64", StringComparison.Ordinal) < 0,
            "installer build helper should no longer drive separate x86/x64 package builds");
    }

    private static void TestSetupScriptDocumentsSinglePublicInstallerAndVerifiedLiveRegistrationFlow()
    {
        string source = ReadProjectSource("Setup.iss");
        string buildScript = ReadProjectSource("build-installer.ps1");

        AssertTrue(source.Contains("OutputBaseFilename=WordToolbox_Setup", StringComparison.Ordinal),
            "setup script should emit a single public installer filename");
        AssertTrue(source.IndexOf("WordToolbox_Setup_x86", StringComparison.Ordinal) < 0,
            "setup script should no longer emit a public x86 package");
        AssertTrue(source.IndexOf("WordToolbox_Setup_x64", StringComparison.Ordinal) < 0,
            "setup script should no longer emit a public x64-only filename");
        AssertTrue(source.Contains("ArchitecturesAllowed=x86compatible", StringComparison.Ordinal),
            "setup script should keep the unified public installer runnable on both 32-bit and 64-bit Windows while Phase 1 expands toward Word x86");
        AssertTrue(source.Contains("ArchitecturesInstallIn64BitMode=x64compatible", StringComparison.Ordinal),
            "setup script should still enter 64-bit install mode when the machine supports it");
        AssertTrue(source.Contains("DefaultDirName={autopf}\\WordToolbox", StringComparison.Ordinal),
            "setup script should resolve the install root through the current platform Program Files view instead of hard-locking to 64-bit Program Files");
        AssertTrue(source.Contains("Source: \"Installer.Core.ps1\"", StringComparison.Ordinal)
            && source.Contains("Source: \"Installer.SupportMatrix.json\"", StringComparison.Ordinal),
            "setup script should bundle the shared installer core and support matrix");
        AssertTrue(source.Contains("BuildLiveRegisterCommand", StringComparison.Ordinal)
            && source.Contains("BuildLiveUnregisterCommand", StringComparison.Ordinal),
            "setup script should build shared-core live register and unregister commands");
        AssertTrue(source.Contains("BuildRegisterSummaryPath", StringComparison.Ordinal)
            && source.Contains("BuildUnregisterSummaryPath", StringComparison.Ordinal),
            "setup script should keep dedicated live summary files for installer-facing diagnostics");
        AssertTrue(source.Contains("BuildInstallerLogDirectory", StringComparison.Ordinal)
            && source.Contains("BuildLiveFailureDiagnostic", StringComparison.Ordinal),
            "setup script should expose installer-facing fallback diagnostics when shared-core summaries are missing");
        AssertTrue(source.Contains("RunSharedCoreLiveRegister", StringComparison.Ordinal)
            && source.Contains("RunSharedCoreLiveUnregister", StringComparison.Ordinal),
            "setup script should verify shared-core live register and unregister through PascalScript");
        AssertTrue(source.Contains("-OutputPath", StringComparison.Ordinal)
            && source.Contains("-SummaryTextPath", StringComparison.Ordinal)
            && source.Contains("-AllowSelfElevation", StringComparison.Ordinal),
            "setup script should require the shared core to write both result and summary files for verification and allow installer-driven self elevation");
        AssertTrue(source.Contains("{localappdata}\\WordTools\\InstallerLogs", StringComparison.Ordinal),
            "setup script should write installer-facing diagnostics to a per-user writable location");
        AssertTrue(source.Contains("Summary: ", StringComparison.Ordinal)
            && source.Contains("Result: ", StringComparison.Ordinal)
            && source.Contains("LogDir: ", StringComparison.Ordinal),
            "setup script should surface log file locations when live registration diagnostics are missing");
        AssertTrue(source.Contains("LoadStringFromFile(BuildRegisterSummaryPath(), ResultMessage)", StringComparison.Ordinal)
            && source.Contains("LoadStringFromFile(BuildUnregisterSummaryPath(), ResultMessage)", StringComparison.Ordinal),
            "setup script should prefer shared-core success summaries when presenting installer completion messages");
        AssertTrue(source.IndexOf("if CurStep = ssInstall then", StringComparison.Ordinal) < 0
            && source.Contains("if CurStep = ssPostInstall then", StringComparison.Ordinal),
            "setup script should delay shared-core live registration until the post-install phase");
        AssertTrue(source.IndexOf("[Run]", StringComparison.Ordinal) < 0
            && source.IndexOf("[UninstallRun]", StringComparison.Ordinal) < 0,
            "setup script should no longer rely on passive Run sections for registration success reporting");
        AssertTrue(source.Contains("64 位 Microsoft Word", StringComparison.Ordinal),
            "setup script should state that only 64-bit Word is supported");
        AssertTrue(source.Contains("32 位 Word", StringComparison.Ordinal)
            && source.Contains("32 位 WPS", StringComparison.Ordinal)
            && source.Contains("64 位 WPS", StringComparison.Ordinal),
            "setup script should explicitly warn about unsupported environments");
        AssertTrue(source.Contains("DetectedHostDetails:", StringComparison.Ordinal),
            "setup script should require host-state details to be present in the probe summary before continuing");
        AssertTrue(buildScript.Contains("WordToolbox_Setup.exe", StringComparison.Ordinal),
            "installer build helper should report a single public installer output");
        AssertTrue(buildScript.IndexOf("ARCH_X86", StringComparison.Ordinal) < 0
            && buildScript.IndexOf("ARCH_X64", StringComparison.Ordinal) < 0,
            "installer build helper should no longer drive separate x86/x64 package builds");
    }

    private static void TestSetupScriptDocumentsSinglePublicInstallerAndVerifiedLiveRegistrationFlowForWordX64AndWpsX86()
    {
        string source = ReadProjectSource("Setup.iss");
        string buildScript = ReadProjectSource("build-installer.ps1");

        AssertTrue(source.Contains("OutputBaseFilename=WordToolbox_Setup", StringComparison.Ordinal),
            "setup script should emit a single public installer filename");
        AssertTrue(source.Contains("ArchitecturesAllowed=x64compatible", StringComparison.Ordinal),
            "setup script should continue targeting 64-bit Windows for the supported host set");
        AssertTrue(source.Contains("BuildProbeCommand", StringComparison.Ordinal)
            && source.Contains("BuildLiveRegisterCommand", StringComparison.Ordinal)
            && source.Contains("BuildLiveUnregisterCommand", StringComparison.Ordinal),
            "setup script should keep dedicated shared-core command builders for probe, register, and unregister");
        AssertTrue(source.Contains(" -RequestedHost Both", StringComparison.Ordinal),
            "setup script should probe and register both supported hosts through the shared installer core");
        AssertTrue(source.Contains("64 位 Microsoft Word", StringComparison.Ordinal)
            && source.Contains("32 位 WPS", StringComparison.Ordinal),
            "setup script should document the supported Word x64 and WPS x86 boundary");
        AssertTrue(source.Contains("32 位 Word", StringComparison.Ordinal)
            && source.Contains("64 位 WPS", StringComparison.Ordinal),
            "setup script should continue warning about the remaining unsupported host combinations");
        AssertTrue(buildScript.Contains("WordToolbox_Setup.exe", StringComparison.Ordinal),
            "installer build helper should report a single public installer output");
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

    private static void TestRegistrationScriptsRejectUnsupportedHostsAndBitnessThroughWrapper()
    {
        string ps1 = ReadProjectSource("RegisterPlugin.ps1");
        string bat = ReadProjectSource("RegisterPlugin.bat");

        AssertTrue(ps1.Contains("64 位 Microsoft Word", StringComparison.Ordinal),
            "PowerShell registration wrapper should explain the only supported host");
        AssertTrue(ps1.Contains("32 位 Word", StringComparison.Ordinal)
            && ps1.Contains("32 位 WPS", StringComparison.Ordinal)
            && ps1.Contains("64 位 WPS", StringComparison.Ordinal),
            "PowerShell registration wrapper should reject unsupported hosts and bitness");
        AssertTrue(ps1.Contains("Installer.Core.ps1", StringComparison.Ordinal),
            "PowerShell registration wrapper should invoke the shared installer core");
        AssertTrue(bat.Contains("RegisterPlugin.ps1", StringComparison.Ordinal),
            "batch registration script should now defer to the PowerShell wrapper");
    }

    private static void TestRegistrationWrapperStopsHardBlockingWordX86BeforeSharedCoreEvaluation()
    {
        string ps1 = ReadProjectSource("RegisterPlugin.ps1");

        AssertTrue(ps1.Contains("Installer.Core.ps1", StringComparison.Ordinal),
            "PowerShell registration wrapper should continue routing through the shared installer core");
        AssertTrue(ps1.IndexOf("if ($resolvedArchitecture -ne \"x64\")", StringComparison.Ordinal) < 0,
            "PowerShell registration wrapper should stop blanket-rejecting Word x86 before the shared core evaluates support state");
        AssertTrue(ps1.Contains("if ($RequestedHost -ne \"Word\")", StringComparison.Ordinal),
            "PowerShell registration wrapper should keep blocking WPS and mixed-host requests until the real WPS activation path is validated");
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

    private static void TestInstallationGuideDocumentsSingleInstallerRealSupportMatrix()
    {
        string guide = ReadProjectSource("INSTALLATION.md");

        AssertTrue(guide.Contains("统一安装包", StringComparison.Ordinal)
            || guide.Contains("单一安装包", StringComparison.Ordinal),
            "installation guide should explain that delivery now uses a single public installer");
        AssertTrue(guide.Contains("64 位 Word", StringComparison.Ordinal),
            "installation guide should mark 64-bit Word as supported");
        AssertTrue(guide.Contains("32 位 Word", StringComparison.Ordinal)
            && guide.Contains("32 位 WPS", StringComparison.Ordinal)
            && guide.Contains("64 位 WPS", StringComparison.Ordinal),
            "installation guide should continue documenting the unsupported host matrix");
        AssertTrue(guide.IndexOf("x86 安装包", StringComparison.Ordinal) < 0
            && guide.IndexOf("x64 安装包", StringComparison.Ordinal) < 0,
            "installation guide should no longer tell users to choose between separate x86 and x64 public packages");
    }

    private static void TestInstallationGuideRemovesOldX86PlaceholderPackageNarrative()
    {
        string guide = ReadProjectSource("INSTALLATION.md");

        AssertTrue(guide.IndexOf("`x86` 安装包仅用于", StringComparison.Ordinal) < 0,
            "installation guide should stop describing a separate x86 placeholder package now that delivery is unified");
    }

    private static void TestReadmePointsToUnifiedInstallerFlow()
    {
        string readme = ReadProjectSource("README.md");

        AssertTrue(readme.Contains("INSTALLATION.md", StringComparison.Ordinal),
            "README should point readers to the dedicated installation guide");
        AssertTrue(readme.Contains("WordToolbox_Setup.exe", StringComparison.Ordinal),
            "README should mention the single public installer package");
        AssertTrue(readme.Contains("64 位 Microsoft Word", StringComparison.Ordinal),
            "README should keep the current support boundary explicit");
        AssertTrue(readme.IndexOf("x64 安装包", StringComparison.Ordinal) < 0
            && readme.IndexOf("x86 安装包", StringComparison.Ordinal) < 0,
            "README should no longer tell users to choose between separate public x86/x64 packages");
    }

    private static void TestRegistrationScriptsAllowWpsX86AndRejectRemainingUnsupportedTargets()
    {
        string ps1 = ReadProjectSource("RegisterPlugin.ps1");
        string bat = ReadProjectSource("RegisterPlugin.bat");

        AssertTrue(ps1.Contains("Installer.Core.ps1", StringComparison.Ordinal),
            "PowerShell registration wrapper should invoke the shared installer core");
        AssertTrue(bat.Contains("RegisterPlugin.ps1", StringComparison.Ordinal),
            "batch registration script should continue deferring to the PowerShell wrapper");
        AssertTrue(ps1.IndexOf("if ($resolvedArchitecture -ne \"x64\")", StringComparison.Ordinal) < 0,
            "PowerShell registration wrapper should no longer blanket-reject every x86 target once WPS x86 is supported");
        AssertTrue(ps1.IndexOf("if ($RequestedHost -ne \"Word\")", StringComparison.Ordinal) < 0,
            "PowerShell registration wrapper should no longer blanket-reject every WPS request once WPS x86 is supported");
        AssertTrue(ps1.Contains("32 位 WPS", StringComparison.Ordinal),
            "PowerShell registration wrapper should mention WPS x86 in the supported boundary");
        AssertTrue(ps1.Contains("32 位 Word", StringComparison.Ordinal)
            && ps1.Contains("64 位 WPS", StringComparison.Ordinal),
            "PowerShell registration wrapper should continue rejecting the remaining unsupported host and bitness combinations");
    }

    private static void TestInstallationGuideDocumentsSingleInstallerWpsX86SupportMatrix()
    {
        string guide = ReadProjectSource("INSTALLATION.md");

        AssertTrue(guide.Contains("统一安装包", StringComparison.Ordinal)
            || guide.Contains("单一安装包", StringComparison.Ordinal),
            "installation guide should explain that delivery now uses a single public installer");
        AssertTrue(guide.Contains("64 位 Word", StringComparison.Ordinal),
            "installation guide should keep 64-bit Word marked as supported");
        AssertTrue(guide.Contains("32 位 WPS：支持", StringComparison.Ordinal),
            "installation guide should promote WPS x86 after live support is validated");
        AssertTrue(guide.Contains("32 位 Word：暂不支持", StringComparison.Ordinal)
            && guide.Contains("64 位 WPS：暂不支持", StringComparison.Ordinal),
            "installation guide should keep only the remaining unsupported targets blocked");
        AssertTrue(guide.IndexOf("x86 安装包", StringComparison.Ordinal) < 0
            && guide.IndexOf("x64 安装包", StringComparison.Ordinal) < 0,
            "installation guide should no longer tell users to choose between separate x86 and x64 public packages");
    }

    private static void TestReadmePointsToUnifiedInstallerFlowWithWpsX86Support()
    {
        string readme = ReadProjectSource("README.md");

        AssertTrue(readme.Contains("INSTALLATION.md", StringComparison.Ordinal),
            "README should point readers to the dedicated installation guide");
        AssertTrue(readme.Contains("WordToolbox_Setup.exe", StringComparison.Ordinal),
            "README should mention the single public installer package");
        AssertTrue(readme.Contains("64 位 Microsoft Word", StringComparison.Ordinal)
            && readme.Contains("32 位 WPS", StringComparison.Ordinal),
            "README should keep the supported live boundary explicit");
        AssertTrue(readme.IndexOf("WPS 和 32 位 Word 仍处于探针/计划阶段", StringComparison.Ordinal) < 0,
            "README should stop describing WPS x86 as probe-only once live support is opened");
        AssertTrue(readme.IndexOf("x64 安装包", StringComparison.Ordinal) < 0
            && readme.IndexOf("x86 安装包", StringComparison.Ordinal) < 0,
            "README should no longer tell users to choose between separate public x86/x64 packages");
    }

    private static void TestSharedCoreRecordsInstallerStateAndUsesItForLiveUninstallFallback()
    {
        string core = ReadProjectSource("Installer.Core.ps1");

        AssertTrue(core.Contains(@"HKLM:\Software\WordTools\InstallerState", StringComparison.Ordinal),
            "shared installer core should persist install-state under the WordTools installer-state registry path");
        AssertTrue(core.Contains("function Save-InstallerState", StringComparison.Ordinal)
            && core.Contains("function Get-InstallerState", StringComparison.Ordinal)
            && core.Contains("function Get-LiveEligibleHostsFromInstallerState", StringComparison.Ordinal)
            && core.Contains("function Clear-InstallerState", StringComparison.Ordinal),
            "shared installer core should define the install-state persistence helpers");
        AssertTrue(core.Contains("Save-InstallerState -ExecutionResult $executionResult", StringComparison.Ordinal),
            "live register execution should record installer state after successful application");
        AssertTrue(core.Contains("$selection = Get-LiveEligibleHostsFromInstallerState", StringComparison.Ordinal),
            "live unregister execution should fall back to recorded installer state when probe selection is empty");
        AssertTrue(core.Contains("Clear-InstallerState", StringComparison.Ordinal),
            "live unregister execution should clear recorded installer state after successful removal");
    }

    private static void TestSharedCorePreparesWordX86InstallerStateAndMessages()
    {
        string core = ReadProjectSource("Installer.Core.ps1");

        AssertTrue(core.Contains("LastIndexOf(\" \")", StringComparison.Ordinal)
            && core.Contains("$hostBitness = $label.Substring($separatorIndex + 1).Trim()", StringComparison.Ordinal),
            "shared installer core should still be able to rehydrate recorded Word x86 installer-state targets through generic label parsing");
        AssertTrue(core.Contains("Only supported Word hosts are currently allowed.", StringComparison.Ordinal),
            "shared installer core should generalize live eligibility errors across supported Word bitnesses");
        AssertTrue(core.IndexOf("validated Word x64 path", StringComparison.Ordinal) < 0,
            "shared installer core should stop hardcoding the Word live enablement narrative to the x64-only path");
    }

    private static void TestSharedCoreRehydratesInstallerStateHostsThroughSupportMatrix()
    {
        string core = ReadProjectSource("Installer.Core.ps1");

        AssertTrue(core.Contains("Get-SupportMatrix", StringComparison.Ordinal),
            "installer-state rehydration should load the support matrix before rebuilding recorded hosts");
        AssertTrue(core.Contains("Get-SupportDecision -SupportMatrix $supportMatrix -HostName \"Word\" -HostBitness \"x64\"", StringComparison.Ordinal)
            || core.Contains("Get-SupportDecision -SupportMatrix $supportMatrix -HostName $hostName -HostBitness $hostBitness", StringComparison.Ordinal),
            "installer-state rehydration should derive host support details from the shared support-decision helper");
        AssertTrue(core.IndexOf("SupportStatus = \"supported\"", StringComparison.Ordinal) < 0,
            "installer-state rehydration should stop hardcoding supported status into reconstructed hosts");
    }

    private static void TestSharedCoreParsesInstallerStateHostLabelsGenerically()
    {
        string core = ReadProjectSource("Installer.Core.ps1");

        AssertTrue(core.Contains("-split \" \"", StringComparison.Ordinal)
            || core.Contains("-split ' '", StringComparison.Ordinal)
            || core.Contains("LastIndexOf(\" \")", StringComparison.Ordinal),
            "installer-state rehydration should parse recorded host labels generically instead of relying on host-specific branches");
        AssertTrue(core.IndexOf("$label -eq \"Word x64\"", StringComparison.Ordinal) < 0
            && core.IndexOf("$label -eq \"Word x86\"", StringComparison.Ordinal) < 0,
            "installer-state rehydration should stop hardcoding Word-specific label branches");
    }

    private static void TestSharedCoreRehydratesInstallerStateExecutionMetadataForDiagnostics()
    {
        string core = ReadProjectSource("Installer.Core.ps1");

        AssertTrue(core.Contains("ExecutionTimestampUtc", StringComparison.Ordinal),
            "installer-state persistence should keep an execution timestamp available for later diagnostics");
        AssertTrue(core.Contains("Get-ItemPropertyValue -Path $registryPath -Name \"ExecutionTimestampUtc\"", StringComparison.Ordinal),
            "installer-state rehydration should read the persisted execution timestamp back from the registry");
        AssertTrue(core.Contains("DiagnosticsBundleId = Get-DiagnosticsBundleId", StringComparison.Ordinal)
            && core.Contains("$state.ExecutionTimestampUtc", StringComparison.Ordinal),
            "recorded host diagnostics identifiers should derive from the persisted execution timestamp rather than unrelated configuration text");
    }

    private static void TestSharedCorePreservesVersionLinesThroughInstallerStateRehydration()
    {
        string core = ReadProjectSource("Installer.Core.ps1");

        AssertTrue(core.Contains("RecordedVersionLines", StringComparison.Ordinal),
            "installer-state persistence should keep recorded host version lines");
        AssertTrue(core.Contains("Get-ItemPropertyValue -Path $registryPath -Name \"RecordedVersionLines\"", StringComparison.Ordinal),
            "installer-state rehydration should read recorded host version lines back from the registry");
        AssertTrue(core.Contains("VersionLine") && core.Contains("$recordedVersionLinesByLabel"),
            "recorded host reconstruction should rehydrate version lines from the persisted installer-state map");
    }

    private static void TestManualVerificationChecklistCoversInstallerStateAndResidualCleanup()
    {
        string checklist = ReadProjectSource(Path.Combine("docs", "installer", "manual-verification-checklist.md"));

        AssertTrue(checklist.Contains("InstallerState", StringComparison.Ordinal),
            "manual verification checklist should require checking the recorded installer state");
        AssertTrue(checklist.Contains(@"HKLM\Software\WordTools\InstallerState", StringComparison.Ordinal),
            "manual verification checklist should document the installer-state registry path");
        AssertTrue(checklist.Contains(@"HKLM\Software\Microsoft\Office\Word\Addins\WordTools.ThisAddIn", StringComparison.Ordinal),
            "manual verification checklist should document the residual Word add-in registry cleanup check");
        AssertTrue(checklist.Contains("ExecutionTimestampUtc", StringComparison.Ordinal),
            "manual verification checklist should document the recorded execution timestamp field");
    }

    private static void TestWpsLiveHandlersAreImplementedBehindDisabledFeatureGate()
    {
        string core = ReadProjectSource("Installer.Core.ps1");

        AssertTrue(core.Contains("function Test-WpsLiveExecutionFeatureEnabled", StringComparison.Ordinal),
            "shared installer core should define a dedicated WPS live-execution feature gate");
        AssertTrue(core.Contains("return $false", StringComparison.Ordinal),
            "WPS live feature gate should remain disabled by default");
        AssertTrue(core.Contains("function Set-WpsAddInRegistry", StringComparison.Ordinal)
            && core.Contains("function Remove-WpsAddInRegistry", StringComparison.Ordinal),
            "shared installer core should define WPS registry write and cleanup helpers");
        AssertTrue(core.Contains("Set-WpsAddInRegistry -RegistryPath $registryTarget", StringComparison.Ordinal)
            && core.Contains("Remove-WpsAddInRegistry -RegistryPath $registryTarget", StringComparison.Ordinal),
            "WPS live handlers should be wired to their dedicated registry helpers");
        AssertTrue(core.Contains("Live WPS registration is not allowed", StringComparison.Ordinal)
            && core.Contains("Live WPS unregistration is not allowed", StringComparison.Ordinal),
            "WPS live handlers should still hard-block execution until the support gate is opened");
    }

    private static void TestManualVerificationChecklistCoversInstallerStateAndWpsResidualCleanup()
    {
        string checklist = ReadProjectSource(Path.Combine("docs", "installer", "manual-verification-checklist.md"));

        AssertTrue(checklist.Contains("InstallerState", StringComparison.Ordinal),
            "manual verification checklist should require checking the recorded installer state");
        AssertTrue(checklist.Contains(@"HKLM\Software\WordTools\InstallerState", StringComparison.Ordinal),
            "manual verification checklist should document the installer-state registry path");
        AssertTrue(checklist.Contains(@"HKLM\Software\Microsoft\Office\Word\Addins\WordTools.ThisAddIn", StringComparison.Ordinal),
            "manual verification checklist should document the residual Word add-in registry cleanup check");
        AssertTrue(checklist.Contains(@"HKCU\Software\Kingsoft\Office\WPS\AddinsWl\WordTools.ThisAddIn", StringComparison.Ordinal),
            "manual verification checklist should document the residual WPS add-in registry cleanup check");
        AssertTrue(checklist.Contains("ExecutionTimestampUtc", StringComparison.Ordinal),
            "manual verification checklist should document the recorded execution timestamp field");
        AssertTrue(checklist.Contains("RecordedVersionLines", StringComparison.Ordinal),
            "manual verification checklist should document the recorded host version-line field");
        AssertTrue(checklist.Contains("DetectedHostDetails", StringComparison.Ordinal),
            "manual verification checklist should require checking the host-state details line in the preview summary");
        AssertTrue(checklist.Contains("DiagnosticsBundleId", StringComparison.Ordinal),
            "manual verification checklist should require checking the host diagnostics bundle identifier");
    }

    private static void TestManualVerificationChecklistRecordsWpsUiFailureAndKeepsWpsUnsupported()
    {
        string checklist = ReadProjectSource(Path.Combine("docs", "installer", "manual-verification-checklist.md"));

        AssertTrue(checklist.Contains("2026-05-24", StringComparison.Ordinal)
            && checklist.Contains("WPS 工具区并未出现插件入口", StringComparison.Ordinal),
            "manual verification checklist should record the observed WPS UI failure from 2026-05-24");
        AssertTrue(checklist.Contains("shared core live register / unregister 成功")
            && checklist.Contains("不能等同于 WPS UI 侧已验证", StringComparison.Ordinal),
            "manual verification checklist should explicitly distinguish shared-core success from WPS UI validation");
        AssertTrue(checklist.Contains("当前版本的正式支持范围只有", StringComparison.Ordinal)
            && checklist.Contains("64 位 Microsoft Word", StringComparison.Ordinal)
            && checklist.Contains("当前仍不支持：", StringComparison.Ordinal)
            && checklist.Contains("32 位 WPS", StringComparison.Ordinal),
            "manual verification checklist should keep WPS x86 out of the formally supported boundary and explicitly list it as unsupported after rollback");
    }

    private static void TestWpsReconEvidenceDocumentsNativePackageMetadata()
    {
        string markdown = ReadProjectSource(Path.Combine("docs", "installer", "evidence", "CurrentMachine-WpsX86-Recon-20260520.md"));

        AssertTrue(markdown.Contains("setupplugin.plg", StringComparison.Ordinal)
            && markdown.Contains("listV3", StringComparison.Ordinal)
            && markdown.Contains("plugin-provider.json", StringComparison.Ordinal)
            && markdown.Contains("runinfo.json", StringComparison.Ordinal),
            "WPS recon evidence should document the native plugin manifests and package metadata that now look stronger than an AddinsWl-only story");
        AssertTrue(markdown.Contains("host", StringComparison.OrdinalIgnoreCase)
            && markdown.Contains("et;wpp;wps;pdf", StringComparison.OrdinalIgnoreCase)
            && markdown.Contains("RunInfoDrivenWebOrJsApi", StringComparison.Ordinal)
            && markdown.Contains("DllAttrNativeModule", StringComparison.Ordinal)
            && markdown.Contains("pool\\win-x64", StringComparison.OrdinalIgnoreCase)
            && markdown.Contains("listV3\\win-x64", StringComparison.OrdinalIgnoreCase)
            && markdown.Contains("present: `%APPDATA%\\Kingsoft\\wps\\addons\\pool\\win-i386`", StringComparison.OrdinalIgnoreCase)
            && markdown.Contains("absent: `%APPDATA%\\Kingsoft\\wps\\addons\\pool\\win-x64`", StringComparison.OrdinalIgnoreCase)
            && markdown.Contains("absent: `%APPDATA%\\Kingsoft\\wps\\addons\\listV3\\win-x64`", StringComparison.OrdinalIgnoreCase)
            && markdown.Contains("entryDll", StringComparison.Ordinal)
            && markdown.Contains("office_type", StringComparison.Ordinal)
            && markdown.Contains("%workingroot%/index.html", StringComparison.Ordinal)
            && markdown.Contains("kplugin", StringComparison.OrdinalIgnoreCase)
            && markdown.Contains("`33` numeric shard directories", StringComparison.OrdinalIgnoreCase)
            && markdown.Contains("`12.1.0.26375\\pluginlist.plg` currently has length `", StringComparison.OrdinalIgnoreCase)
            && markdown.Contains("`12.1.0.26375\\515615`", StringComparison.OrdinalIgnoreCase)
            && markdown.Contains("`ids` length `697` bytes", StringComparison.OrdinalIgnoreCase)
            && markdown.Contains("`data` length `31769` bytes", StringComparison.OrdinalIgnoreCase)
            && markdown.Contains("did **not** appear inside the sampled `pluginlist.plg` as a plain ASCII or UTF-16LE literal", StringComparison.Ordinal)
            && markdown.Contains("`UInt32` little-endian byte pattern `1F-DE-07-00`", StringComparison.Ordinal)
            && markdown.Contains("did not expose `kdocerjsapi20`, `kwpsaiwordtool`, or `kdocsword` as plain ASCII literals", StringComparison.Ordinal)
            && markdown.Contains("`picture_resourceshop_split`, `kdocerjsapi20.dll`, `CreateSplitAppWidget`, or `proxyFrame`", StringComparison.Ordinal)
            && markdown.Contains("or UTF-16LE literals", StringComparison.Ordinal)
            && markdown.Contains("GetExtensionJsApiObj", StringComparison.Ordinal)
            && markdown.Contains("DllRegisterServer", StringComparison.Ordinal),
            "WPS recon evidence should record field-level native metadata clues from setupplugin host scope, mapped runtime-shape classification, runinfo/config/run.ini, binary index headers, negative literal-search findings, and sampled native DLL export evidence");
        AssertTrue(markdown.Contains(@"HKCU\Software\Kingsoft\Office\WPS\AddinsWl", StringComparison.Ordinal)
            && markdown.Contains("version gates", StringComparison.OrdinalIgnoreCase)
            && markdown.Contains("not like a table of resolvable live ProgIDs", StringComparison.Ordinal),
            "WPS recon evidence should keep the AddinsWl write contract explicitly unvalidated even after the newer native package signals are recorded");
    }

    private static void TestSupportMatrixPromotesWordX64AndKeepsOthersPlanned()
    {
        string matrix = ReadProjectSource("Installer.SupportMatrix.json");

        AssertTrue(matrix.Contains("\"host\": \"Word\"") && matrix.Contains("\"bitness\": \"x86\""),
            "support matrix should include Word x86 as a planned target");
        AssertTrue(matrix.Contains("\"host\": \"Word\"") && matrix.Contains("\"bitness\": \"x64\""),
            "support matrix should include Word x64 as a validated target");
        AssertTrue(matrix.Contains("\"host\": \"WPS\"") && matrix.Contains("\"bitness\": \"x86\""),
            "support matrix should include WPS x86 as a tracked target");
        AssertTrue(matrix.Contains("\"host\": \"WPS\"") && matrix.Contains("\"bitness\": \"x64\""),
            "support matrix should include WPS x64 as a planned target");
        AssertTrue(matrix.Contains("\"host\": \"WPS\"") && matrix.Contains("\"bitness\": \"x86\"") && matrix.Contains("\"status\": \"planned\""),
            "support matrix should roll WPS x86 back to planned when the UI validation fails");
        AssertTrue(matrix.Contains("\"host\": \"WPS\"") && matrix.Contains("\"bitness\": \"x64\"") && matrix.Contains("\"status\": \"planned\""),
            "support matrix should keep WPS x64 in planned status");
        AssertTrue(matrix.Contains("\"host\": \"Word\"") && matrix.Contains("\"bitness\": \"x86\"") && matrix.Contains("\"status\": \"planned\""),
            "support matrix should keep Word x86 in planned status");
    }

    private static void TestSupportMatrixDeclaresValidationStagesAndActivationRoutes()
    {
        string matrixPath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Installer.SupportMatrix.json"));
        string matrixText = File.ReadAllText(matrixPath);
        using JsonDocument matrix = JsonDocument.Parse(matrixText);

        AssertTrue(matrixText.Contains("\"validationStage\""),
            "support matrix should declare a validation stage for each tracked host target");
        AssertTrue(matrixText.Contains("\"activationRoute\""),
            "support matrix should declare an activation route for each tracked host target");

        JsonElement targets = matrix.RootElement.GetProperty("targets");
        AssertTrue(targets.EnumerateArray().Any(entry =>
            string.Equals(entry.GetProperty("host").GetString(), "Word", StringComparison.Ordinal)
            && string.Equals(entry.GetProperty("bitness").GetString(), "x64", StringComparison.Ordinal)
            && string.Equals(entry.GetProperty("validationStage").GetString(), "formal p0 passed", StringComparison.Ordinal)
            && string.Equals(entry.GetProperty("activationRoute").GetString(), "WordComAddin", StringComparison.Ordinal)),
            "support matrix should mark Word x64 as a formal P0-passed WordComAddin target");
        AssertTrue(targets.EnumerateArray().Any(entry =>
            string.Equals(entry.GetProperty("host").GetString(), "Word", StringComparison.Ordinal)
            && string.Equals(entry.GetProperty("bitness").GetString(), "x86", StringComparison.Ordinal)
            && string.Equals(entry.GetProperty("validationStage").GetString(), "probe only", StringComparison.Ordinal)
            && string.Equals(entry.GetProperty("activationRoute").GetString(), "WordComAddin", StringComparison.Ordinal)),
            "support matrix should keep Word x86 in the probe-only WordComAddin stage until real evidence exists");
        AssertTrue(targets.EnumerateArray().Any(entry =>
            string.Equals(entry.GetProperty("host").GetString(), "WPS", StringComparison.Ordinal)
            && string.Equals(entry.GetProperty("bitness").GetString(), "x86", StringComparison.Ordinal)
            && string.Equals(entry.GetProperty("validationStage").GetString(), "experimenting", StringComparison.Ordinal)
            && string.Equals(entry.GetProperty("activationRoute").GetString(), "WpsNativePlugin", StringComparison.Ordinal)),
            "support matrix should mark WPS x86 as a WpsNativePlugin target currently in the experimenting stage");
    }

    private static void TestInvokeWpsAddinsWlExperimentDefined()
    {
        string core = ReadProjectSource("Installer.Core.ps1");

        AssertTrue(core.Contains("function Invoke-WpsAddinsWlExperiment"),
            "Installer.Core.ps1 should define Invoke-WpsAddinsWlExperiment function");
        AssertTrue(core.Contains("$ProgId")
            && core.Contains("$ValuePayload")
            && core.Contains("$ExperimentId")
            && core.Contains("$EvidenceDir"),
            "Invoke-WpsAddinsWlExperiment should expose the four mandatory parameters: ProgId, ValuePayload, ExperimentId, EvidenceDir");
        AssertTrue(core.Contains("[switch]$Restore"),
            "Invoke-WpsAddinsWlExperiment should expose a Restore switch parameter");
        AssertTrue(core.Contains("reg export")
            && core.Contains("reg import"),
            "Invoke-WpsAddinsWlExperiment should use reg export/import for backup and restore");
        AssertTrue(core.Contains("Set-ItemProperty"),
            "Invoke-WpsAddinsWlExperiment should use Set-ItemProperty to write into AddinsWl");
        AssertTrue(core.Contains("PreExisting"),
            "Invoke-WpsAddinsWlExperiment should detect pre-existing AddinsWl entries before writing");
    }

    private static void TestInvokeWpsAddinsWlExperimentBackupFlow()
    {
        string core = ReadProjectSource("Installer.Core.ps1");

        AssertTrue(core.Contains("Test-Path -LiteralPath $backupPath -PathType Leaf"),
            "Invoke-WpsAddinsWlExperiment should verify the backup .reg file exists before proceeding");
        AssertTrue(core.Contains("Backup failed"),
            "Invoke-WpsAddinsWlExperiment should report a backup-failure error when reg export produces no output");
        AssertTrue(core.Contains("AddinsWlPreTotal")
            && core.Contains("AddinsWlPostTotal"),
            "Invoke-WpsAddinsWlExperiment should track entry counts before and after the write");
        AssertTrue(core.Contains("PostRestoreEntryCount"),
            "Invoke-WpsAddinsWlExperiment restore mode should report the entry count after restore");
    }

    private static void TestSupportMatrixHasUiExperimentState()
    {
        string matrixPath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Installer.SupportMatrix.json"));
        string matrixText = File.ReadAllText(matrixPath);
        using JsonDocument matrix = JsonDocument.Parse(matrixText);

        AssertTrue(matrixText.Contains("\"ui-experiment\""),
            "support matrix should declare the new ui-experiment status for hosts currently undergoing incremental experiment-chain validation");

        JsonElement targets = matrix.RootElement.GetProperty("targets");
        AssertTrue(targets.EnumerateArray().Any(entry =>
            string.Equals(entry.GetProperty("host").GetString(), "WPS", StringComparison.Ordinal)
            && string.Equals(entry.GetProperty("bitness").GetString(), "x86", StringComparison.Ordinal)
            && string.Equals(entry.GetProperty("status").GetString(), "ui-experiment", StringComparison.Ordinal)
            && string.Equals(entry.GetProperty("validationStage").GetString(), "experimenting", StringComparison.Ordinal)),
            "support matrix should mark WPS x86 as ui-experiment with an experimenting validation stage");
    }

    private static void TestProbeCoreScaffoldExistsWithoutChangingRegistrationPath()
    {
        string core = ReadProjectSource("Installer.Core.ps1");

        AssertTrue(core.Contains("-Mode Probe"),
            "probe phase should add a dedicated probe mode to the shared installer core");
        AssertTrue(core.Contains("function Get-HostInventory"),
            "probe phase should define a host inventory function");
        AssertTrue(core.Contains("function Get-InstalledWordCandidates"),
            "probe phase should define a Word candidate detector");
        AssertTrue(core.Contains("function Get-InstalledWpsCandidates"),
            "probe phase should define a WPS candidate detector");
        AssertTrue(core.Contains("function Get-ExecutableBitness"),
            "probe phase should define an executable bitness detector");
    }

    private static void TestProbePhasePreservesCurrentWordOnlyRegistrationPath()
    {
        string ps1 = ReadProjectSource("RegisterPlugin.ps1");

        AssertTrue(ps1.IndexOf("Installer.Core.ps1", StringComparison.Ordinal) < 0,
            "probe phase must not reroute the current registration entrypoint yet");
        AssertTrue(ps1.Contains("仅支持 64 位 Microsoft Word"),
            "probe phase must preserve the current user-facing Word-only support warning");
        AssertTrue(ps1.Contains("WordTools.ThisAddIn"),
            "probe phase must preserve the current Word add-in ProgId path");
        AssertTrue(ps1.Contains("Microsoft\\Office\\Word\\Addins"),
            "probe phase must preserve the current Word add-in registry path");
    }

    private static void TestUnifiedCoreExposesRegisterAndUnregisterPreviewModesWithoutReroutingEntryPoint()
    {
        string core = ReadProjectSource("Installer.Core.ps1");
        string ps1 = ReadProjectSource("RegisterPlugin.ps1");

        AssertTrue(core.Contains("[ValidateSet(\"Probe\", \"Plan\", \"Register\", \"Unregister\")]"),
            "unified installer core should declare register and unregister modes alongside probe and plan");
        AssertTrue(core.Contains("function Get-RegisterPreviewPlan"),
            "unified installer core should define a register preview planner");
        AssertTrue(core.Contains("function Get-UnregisterPreviewPlan"),
            "unified installer core should define an unregister preview planner");
        AssertTrue(core.Contains("function Get-RegisterPreviewTargetForWordHost"),
            "unified installer core should define a Word-specific register preview target builder");
        AssertTrue(core.Contains("function Get-RegisterPreviewTargetForWpsHost"),
            "unified installer core should define a WPS-specific register preview target builder");
        AssertTrue(core.Contains("function Get-UnregisterPreviewTargetForWordHost"),
            "unified installer core should define a Word-specific unregister preview target builder");
        AssertTrue(core.Contains("function Get-UnregisterPreviewTargetForWpsHost"),
            "unified installer core should define a WPS-specific unregister preview target builder");
        AssertTrue(core.Contains("function Invoke-RegAsm32"),
            "unified installer core should define a 32-bit RegAsm helper skeleton");
        AssertTrue(core.Contains("function Invoke-RegAsm64"),
            "unified installer core should define a 64-bit RegAsm helper skeleton");
        AssertTrue(core.Contains("function Register-WordHost"),
            "unified installer core should define a Word-specific live registration skeleton");
        AssertTrue(core.Contains("function Register-WpsHost"),
            "unified installer core should define a WPS-specific live registration skeleton");
        AssertTrue(core.Contains("function Unregister-WordHost"),
            "unified installer core should define a Word-specific live unregistration skeleton");
        AssertTrue(core.Contains("function Unregister-WpsHost"),
            "unified installer core should define a WPS-specific live unregistration skeleton");
        AssertTrue(core.Contains("PreviewOnly"),
            "register orchestration should stay preview-only until the live entrypoint is intentionally rerouted");
        AssertTrue(ps1.IndexOf("Installer.Core.ps1", StringComparison.Ordinal) < 0,
            "adding preview register modes must not reroute the current registration entrypoint yet");

        string scriptPath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Installer.Core.ps1"));
        AssertPreviewMode(scriptPath, "Register", "RegisterPlan", "RegisterPreviewSummary", new[] { "RegisterComCodebase", "WriteWordAddInRegistry" });
        AssertPreviewMode(scriptPath, "Unregister", "UnregisterPlan", "UnregisterPreviewSummary", new[] { "UnregisterComCodebase", "RemoveWordAddInRegistry" });
    }

    private static void TestReroutedRegistrationEntryPointPreservesWordOnlySupportGuardrails()
    {
        string ps1 = ReadProjectSource("RegisterPlugin.ps1");
        string bat = ReadProjectSource("RegisterPlugin.bat");

        AssertTrue(ps1.Contains("Installer.Core.ps1", StringComparison.Ordinal),
            "the PowerShell registration entrypoint should reroute through the shared installer core");
        AssertTrue(ps1.Contains("-ExecutionIntent", StringComparison.Ordinal)
            && ps1.Contains("Live", StringComparison.Ordinal),
            "the PowerShell wrapper should invoke the shared installer core in live mode");
        AssertTrue(ps1.Contains("64 位 Microsoft Word", StringComparison.Ordinal),
            "the PowerShell wrapper should preserve the current Word-only support message");
        AssertTrue(ps1.Contains("32 位 Word", StringComparison.Ordinal)
            && ps1.Contains("32 位 WPS", StringComparison.Ordinal)
            && ps1.Contains("64 位 WPS", StringComparison.Ordinal),
            "the PowerShell wrapper should preserve the unsupported-host guidance");
        AssertTrue(ps1.Contains("WordTools.ThisAddIn", StringComparison.Ordinal),
            "the PowerShell wrapper should preserve the current Word add-in ProgId");
        AssertTrue(ps1.Contains(@"Microsoft\Office\Word\Addins", StringComparison.Ordinal),
            "the PowerShell wrapper should preserve the current Word add-in registry path");
        AssertTrue(bat.Contains("RegisterPlugin.ps1", StringComparison.Ordinal),
            "the batch entrypoint should now forward to the PowerShell wrapper");
    }

    private static void TestUnifiedCoreExposesRegisterAndUnregisterPreviewModesAlongsideReroutedLiveEntrypoint()
    {
        string core = ReadProjectSource("Installer.Core.ps1");
        string ps1 = ReadProjectSource("RegisterPlugin.ps1");
        string bat = ReadProjectSource("RegisterPlugin.bat");

        AssertTrue(core.Contains("[ValidateSet(\"Probe\", \"Plan\", \"Register\", \"Unregister\")]"),
            "unified installer core should declare register and unregister modes alongside probe and plan");
        AssertTrue(core.Contains("[ValidateSet(\"PreviewOnly\", \"Live\")]"),
            "unified installer core should declare an execution-intent gate for preview versus live execution");
        AssertTrue(core.Contains("[ValidateSet(\"Auto\", \"x86\", \"x64\")]"),
            "unified installer core should accept the architecture selector needed by the rerouted entrypoint");
        AssertTrue(core.Contains("[ValidateSet(\"Word\", \"WPS\", \"Both\")]"),
            "unified installer core should accept the host selector needed by the rerouted entrypoint");
        AssertTrue(core.Contains("function Get-RegisterPreviewPlan"),
            "unified installer core should define a register preview planner");
        AssertTrue(core.Contains("function Get-UnregisterPreviewPlan"),
            "unified installer core should define an unregister preview planner");
        AssertTrue(core.Contains("function Invoke-RegAsm32"),
            "unified installer core should define a 32-bit RegAsm helper");
        AssertTrue(core.Contains("function Invoke-RegAsm64"),
            "unified installer core should define a 64-bit RegAsm helper");
        AssertTrue(core.Contains("function Register-WordHost"),
            "unified installer core should define a Word-specific live registration handler");
        AssertTrue(core.Contains("function Register-WpsHost"),
            "unified installer core should define a WPS-specific live registration handler");
        AssertTrue(core.Contains("function Unregister-WordHost"),
            "unified installer core should define a Word-specific live unregistration handler");
        AssertTrue(core.Contains("function Unregister-WpsHost"),
            "unified installer core should define a WPS-specific live unregistration handler");
        AssertTrue(core.Contains("ExecutionIntent", StringComparison.Ordinal),
            "unified installer core should expose the execution intent to distinguish preview from live runs");
        AssertTrue(ps1.Contains("Installer.Core.ps1", StringComparison.Ordinal),
            "the PowerShell live registration entrypoint should reroute through the shared core");
        AssertTrue(bat.Contains("RegisterPlugin.ps1", StringComparison.Ordinal),
            "the batch live registration entrypoint should reroute through the PowerShell wrapper");

        string scriptPath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Installer.Core.ps1"));
        AssertPreviewModeAfterReroute(scriptPath, "Register", "RegisterPlan", "RegisterPreviewSummary", new[] { "RegisterComCodebase", "WriteWordAddInRegistry" });
        AssertPreviewModeAfterReroute(scriptPath, "Unregister", "UnregisterPlan", "UnregisterPreviewSummary", new[] { "UnregisterComCodebase", "RemoveWordAddInRegistry" });
    }

    private static void AssertPreviewModeAfterReroute(string scriptPath, string mode, string payloadKey, string summaryKey, string[] requiredActions)
    {
        var startInfo = new System.Diagnostics.ProcessStartInfo
        {
            FileName = "powershell",
            Arguments = "-NoProfile -ExecutionPolicy Bypass -File \"" + scriptPath + "\" -Mode " + mode + " -EvidenceLabel \"CurrentMachine-Word64-20260517\"",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        using var process = System.Diagnostics.Process.Start(startInfo);
        string standardOutput = process.StandardOutput.ReadToEnd();
        string standardError = process.StandardError.ReadToEnd();
        process.WaitForExit();

        AssertEqual(0, process.ExitCode, mode + " mode should complete successfully as a preview. stderr: " + standardError);

        using JsonDocument json = JsonDocument.Parse(standardOutput);
        JsonElement previewPlan = json.RootElement.GetProperty(payloadKey);
        JsonElement previewSummary = previewPlan.GetProperty(summaryKey);

        AssertEqual("PreviewOnly", previewSummary.GetProperty("ExecutionMode").GetString(),
            mode + " summary should remain explicit about preview-only execution");
        AssertEqual(1, previewSummary.GetProperty("DetectedTargetCount").GetInt32(),
            mode + " summary should reflect the single detected Word x64 target on the current machine");
        AssertEqual(1, previewSummary.GetProperty("PreviewableTargetCount").GetInt32(),
            mode + " summary should mark the supported Word x64 host as previewable");

        string[] requiredRegAsmModes = previewSummary.GetProperty("RequiredRegAsmModes")
            .EnumerateArray()
            .Select(static item => item.GetString())
            .Where(static item => !string.IsNullOrWhiteSpace(item))
            .ToArray();
        AssertTrue(requiredRegAsmModes.SequenceEqual(new[] { "x64" }, StringComparer.Ordinal),
            mode + " summary should require only x64 RegAsm on the current machine");

        string[] actionUnion = previewSummary.GetProperty("PlannedActionUnion")
            .EnumerateArray()
            .Select(static item => item.GetString())
            .Where(static item => !string.IsNullOrWhiteSpace(item))
            .ToArray();
        foreach (string action in requiredActions)
        {
            AssertTrue(actionUnion.Contains(action, StringComparer.Ordinal),
                mode + " summary should include " + action + " in the action union");
        }

        JsonElement liveReadinessSummary = previewPlan.GetProperty("LiveReadinessSummary");
        AssertEqual(1, liveReadinessSummary.GetProperty("LiveReadyHostCount").GetInt32(),
            mode + " live-readiness summary should mark the supported Word x64 host as live-ready after entrypoint reroute");
        AssertEqual(0, liveReadinessSummary.GetProperty("PreviewOnlyHostCount").GetInt32(),
            mode + " live-readiness summary should no longer classify the current supported host as preview-only");
        AssertTrue(liveReadinessSummary.GetProperty("LiveReadyHostLabels")
            .EnumerateArray()
            .Select(static item => item.GetString())
            .Any(static item => string.Equals(item, "Word x64", StringComparison.Ordinal)),
            mode + " live-readiness summary should expose the current Word x64 host as live-ready");

        JsonElement manifestItem = previewPlan.GetProperty("OperationManifest")[0];
        AssertEqual(true, manifestItem.GetProperty("LiveExecutionAllowed").GetBoolean(),
            mode + " operation manifest should show that live execution is now allowed for the current Word x64 host");

        JsonElement installerHandoffSummary = previewPlan.GetProperty("InstallerHandoffSummary");
        AssertEqual(1, installerHandoffSummary.GetProperty("SupportedHostCount").GetInt32(),
            mode + " installer handoff summary should treat the current Word x64 host as supported");

        JsonElement liveEntrypointStatus = previewPlan.GetProperty("LiveEntrypointStatus");
        AssertTrue(liveEntrypointStatus.GetProperty("ReroutedEntrypoints")
            .EnumerateArray()
            .Select(static item => item.GetString())
            .Any(static item => string.Equals(item, "RegisterPlugin.ps1", StringComparison.Ordinal)),
            mode + " preview plan should document that the PowerShell registration entrypoint is rerouted");
        AssertTrue(liveEntrypointStatus.GetProperty("ReroutedEntrypoints")
            .EnumerateArray()
            .Select(static item => item.GetString())
            .Any(static item => string.Equals(item, "Setup.iss", StringComparison.Ordinal)),
            mode + " preview plan should document that the installer entrypoint is rerouted");
        AssertEqual(0, liveEntrypointStatus.GetProperty("PendingEntrypoints").GetArrayLength(),
            mode + " preview plan should no longer report pending live entrypoints for the current supported flow");
        AssertEqual(true, liveEntrypointStatus.GetProperty("SharedCoreOwnsLiveRegistration").GetBoolean(),
            mode + " preview plan should confirm that the shared core now owns live registration for the current supported flow");
        AssertTrue((liveEntrypointStatus.GetProperty("MigrationDecision").GetString() ?? string.Empty)
            .Contains("Setup.iss", StringComparison.Ordinal)
            && (liveEntrypointStatus.GetProperty("MigrationDecision").GetString() ?? string.Empty)
                .Contains("delegate live registration", StringComparison.OrdinalIgnoreCase),
            mode + " preview plan should clearly state that the installer and script entrypoints delegate live registration to the shared core");

        JsonElement migrationChecklist = previewPlan.GetProperty("MigrationChecklist");
        AssertEqual(true, migrationChecklist.GetProperty("ReadyToRerouteLiveEntrypoints").GetBoolean(),
            mode + " migration checklist should mark live-entrypoint reroute as complete for the current supported flow");
        AssertEqual(0, migrationChecklist.GetProperty("BlockingItems").GetArrayLength(),
            mode + " migration checklist should no longer report blocking items after the installer path is rerouted");
        AssertTrue(migrationChecklist.GetProperty("DeferredSupportTargets")
            .EnumerateArray()
            .Select(static item => item.GetString())
            .Any(static item => item?.Contains("WPS x64", StringComparison.Ordinal) == true),
            mode + " migration checklist should keep unvalidated support-expansion targets visible as deferred items");
        AssertTrue((migrationChecklist.GetProperty("OverallDecision").GetString() ?? string.Empty)
            .Contains("support expansion", StringComparison.OrdinalIgnoreCase),
            mode + " migration checklist should distinguish completed live reroute from future support expansion");

        JsonElement installerPreviewReport = previewPlan.GetProperty("InstallerPreviewReport");
        AssertEqual(true, installerPreviewReport.GetProperty("ReadyToRerouteLiveEntrypoints").GetBoolean(),
            mode + " installer preview report should expose that live entrypoint reroute is complete for the current supported flow");

        JsonElement firstTarget = previewPlan.GetProperty("Targets")[0];
        JsonElement handlerPreview = firstTarget.GetProperty("HandlerPreview");
        JsonElement hostRuleSummary = handlerPreview.GetProperty("HostRuleSummary");
        AssertEqual(true, hostRuleSummary.GetProperty("LiveExecutionAllowed").GetBoolean(),
            mode + " host rule summary should show that the current Word x64 host is eligible for live execution");
        AssertTrue(hostRuleSummary.GetProperty("EnablementCondition").GetString()?.Contains("shared core can execute live", StringComparison.OrdinalIgnoreCase) == true,
            mode + " host rule summary should document that the shared core can now execute live Word x64 registration");

        if (string.Equals(mode, "Register", StringComparison.Ordinal))
        {
            AssertEqual("Register-WordHost", manifestItem.GetProperty("DispatchHandler").GetString(),
                "register operation manifest should point at the Word register handler");
            AssertEqual("Invoke-RegAsm64", manifestItem.GetProperty("RegAsmInvoker").GetString(),
                "register operation manifest should capture the 64-bit RegAsm invoker");
        }
        else
        {
            AssertEqual("Unregister-WordHost", manifestItem.GetProperty("DispatchHandler").GetString(),
                "unregister operation manifest should point at the Word unregister handler");
            AssertEqual("Invoke-RegAsm64", manifestItem.GetProperty("RegAsmInvoker").GetString(),
                "unregister operation manifest should capture the 64-bit RegAsm invoker");
        }
    }

    private static void AssertPreviewMode(string scriptPath, string mode, string payloadKey, string summaryKey, string[] requiredActions)
    {
        var startInfo = new System.Diagnostics.ProcessStartInfo
        {
            FileName = "powershell",
            Arguments = "-NoProfile -ExecutionPolicy Bypass -File \"" + scriptPath + "\" -Mode " + mode + " -EvidenceLabel \"CurrentMachine-Word64-20260517\"",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        using var process = System.Diagnostics.Process.Start(startInfo);
        string standardOutput = process.StandardOutput.ReadToEnd();
        string standardError = process.StandardError.ReadToEnd();
        process.WaitForExit();

        AssertEqual(0, process.ExitCode, mode + " mode should complete successfully as a preview. stderr: " + standardError);
        AssertTrue(standardOutput.Contains("\"" + payloadKey + "\""),
            mode + " mode should emit the expected preview payload");

        using JsonDocument json = JsonDocument.Parse(standardOutput);
        JsonElement previewPlan = json.RootElement.GetProperty(payloadKey);
        JsonElement previewSummary = previewPlan.GetProperty(summaryKey);

        AssertTrue(standardOutput.Contains("\"ExecutionMode\":  \"PreviewOnly\"")
            || standardOutput.Contains("\"ExecutionMode\":\"PreviewOnly\""),
            mode + " mode should stay preview-only until the live registration entrypoint is explicitly rerouted");
        AssertEqual("PreviewOnly", previewSummary.GetProperty("ExecutionMode").GetString(),
            mode + " summary should remain explicit about preview-only execution");
        AssertEqual(1, previewSummary.GetProperty("DetectedTargetCount").GetInt32(),
            mode + " summary should reflect the single detected Word x64 target on the current machine");
        AssertEqual(1, previewSummary.GetProperty("PreviewableTargetCount").GetInt32(),
            mode + " summary should mark the supported Word x64 host as previewable");
        AssertEqual(0, previewSummary.GetProperty("SkippedTargetCount").GetInt32(),
            mode + " summary should not skip the supported Word x64 host on the current machine");

        string[] requiredRegAsmModes = previewSummary.GetProperty("RequiredRegAsmModes")
            .EnumerateArray()
            .Select(static item => item.GetString())
            .Where(static item => !string.IsNullOrWhiteSpace(item))
            .ToArray();
        AssertTrue(requiredRegAsmModes.SequenceEqual(new[] { "x64" }, StringComparer.Ordinal),
            mode + " summary should require only x64 RegAsm on the current machine");

        string[] registryWrites = previewSummary.GetProperty("RegistryWrites")
            .EnumerateArray()
            .Select(static item => item.GetString())
            .Where(static item => !string.IsNullOrWhiteSpace(item))
            .ToArray();
        AssertTrue(registryWrites.Any(static item => item.Contains(@"Microsoft\Office\Word\Addins\WordTools.ThisAddIn", StringComparison.Ordinal)),
            mode + " summary should include the Word add-in registry path");

        string[] actionUnion = previewSummary.GetProperty("PlannedActionUnion")
            .EnumerateArray()
            .Select(static item => item.GetString())
            .Where(static item => !string.IsNullOrWhiteSpace(item))
            .ToArray();
        foreach (string action in requiredActions)
        {
            AssertTrue(actionUnion.Contains(action, StringComparer.Ordinal),
                mode + " summary should include " + action + " in the action union");
        }

        string overallDecision = previewSummary.GetProperty("OverallDecision").GetString() ?? string.Empty;
        AssertTrue(overallDecision.Contains("Preview-only", StringComparison.OrdinalIgnoreCase)
            && overallDecision.Contains("without executing any live", StringComparison.OrdinalIgnoreCase),
            mode + " summary should explicitly state that no live action is executed");

        JsonElement hostRuleSummaries = previewPlan.GetProperty("HostRuleSummaries");
        AssertEqual(1, hostRuleSummaries.GetArrayLength(),
            mode + " preview plan should expose a single grouped host rule summary on the current machine");
        JsonElement topLevelHostRuleSummary = hostRuleSummaries[0];
        AssertEqual("PreviewOnly", topLevelHostRuleSummary.GetProperty("ExecutionMode").GetString(),
            mode + " preview plan should keep grouped host rule summaries in preview-only mode");
        AssertEqual("Word", topLevelHostRuleSummary.GetProperty("HostName").GetString(),
            mode + " preview plan should group the current machine under the Word host rule summary");
        AssertEqual(mode, topLevelHostRuleSummary.GetProperty("Operation").GetString(),
            mode + " preview plan should preserve the operation on the grouped host rule summary");
        AssertEqual(true, topLevelHostRuleSummary.GetProperty("LiveExecutionAllowed").GetBoolean(),
            mode + " grouped host rule summary should expose the current Word x64 host as live-eligible");

        JsonElement liveReadinessSummary = previewPlan.GetProperty("LiveReadinessSummary");
        AssertEqual("PreviewOnly", liveReadinessSummary.GetProperty("ExecutionMode").GetString(),
            mode + " preview plan should expose a preview-only live-readiness summary");
        AssertEqual(1, liveReadinessSummary.GetProperty("DetectedHostCount").GetInt32(),
            mode + " live-readiness summary should count the single detected host on the current machine");
        AssertEqual(1, liveReadinessSummary.GetProperty("LiveReadyHostCount").GetInt32(),
            mode + " live-readiness summary should mark the current supported host as live-ready");
        AssertEqual(0, liveReadinessSummary.GetProperty("PreviewOnlyHostCount").GetInt32(),
            mode + " live-readiness summary should no longer classify the current supported host as preview-only");
        AssertEqual(0, liveReadinessSummary.GetProperty("ProbePendingHostCount").GetInt32(),
            mode + " live-readiness summary should not mark the current supported Word host as probe-pending");
        AssertTrue(liveReadinessSummary.GetProperty("LiveReadyHostLabels")
            .EnumerateArray()
            .Select(static item => item.GetString())
            .Any(static item => string.Equals(item, "Word x64", StringComparison.Ordinal)),
            mode + " live-readiness summary should expose the current Word x64 host as live-ready");
        AssertEqual(0, liveReadinessSummary.GetProperty("PreviewOnlyHostLabels").GetArrayLength(),
            mode + " live-readiness summary should expose no preview-only hosts on the current machine");
        AssertEqual(0, liveReadinessSummary.GetProperty("ProbePendingHostLabels").GetArrayLength(),
            mode + " live-readiness summary should expose no probe-pending detected hosts on the current machine");
        AssertTrue((liveReadinessSummary.GetProperty("OverallDecision").GetString() ?? string.Empty)
            .Contains("live readiness", StringComparison.OrdinalIgnoreCase)
            || (liveReadinessSummary.GetProperty("OverallDecision").GetString() ?? string.Empty)
                .Contains("live-ready", StringComparison.OrdinalIgnoreCase),
            mode + " live-readiness summary should explicitly call out the live-ready state");

        JsonElement operationManifest = previewPlan.GetProperty("OperationManifest");
        AssertEqual(1, operationManifest.GetArrayLength(),
            mode + " preview plan should expose a single manifest operation on the current machine");
        JsonElement manifestItem = operationManifest[0];
        AssertEqual("Word x64", manifestItem.GetProperty("HostLabel").GetString(),
            mode + " operation manifest should preserve the detected host label");
        AssertEqual(mode, manifestItem.GetProperty("Operation").GetString(),
            mode + " operation manifest should preserve the requested operation");
        AssertEqual("PreviewOnly", manifestItem.GetProperty("ExecutionMode").GetString(),
            mode + " operation manifest should stay preview-only");
        AssertEqual(true, manifestItem.GetProperty("LiveExecutionAllowed").GetBoolean(),
            mode + " operation manifest should expose the current Word x64 host as live-eligible");
        AssertTrue(manifestItem.GetProperty("RegistryTargets")
            .EnumerateArray()
            .Select(static item => item.GetString())
            .Any(static item => string.Equals(item, @"HKLM:\Software\Microsoft\Office\Word\Addins\WordTools.ThisAddIn", StringComparison.Ordinal)),
            mode + " operation manifest should surface the Word add-in registry target");

        JsonElement installerHandoffSummary = previewPlan.GetProperty("InstallerHandoffSummary");
        AssertEqual("PreviewOnly", installerHandoffSummary.GetProperty("ExecutionMode").GetString(),
            mode + " preview plan should expose a preview-only installer handoff summary");
        AssertEqual(1, installerHandoffSummary.GetProperty("DetectedHostCount").GetInt32(),
            mode + " installer handoff summary should count the single detected host on the current machine");
        AssertEqual(1, installerHandoffSummary.GetProperty("SupportedHostCount").GetInt32(),
            mode + " installer handoff summary should treat the current Word x64 host as supported");
        AssertEqual(0, installerHandoffSummary.GetProperty("UnsupportedHostCount").GetInt32(),
            mode + " installer handoff summary should not report unsupported detected hosts on the current machine");
        AssertEqual(1, installerHandoffSummary.GetProperty("PreviewActionCount").GetInt32(),
            mode + " installer handoff summary should point to a single preview action on the current machine");
        AssertTrue(installerHandoffSummary.GetProperty("SupportedHosts")
            .EnumerateArray()
            .Select(static item => item.GetString())
            .Any(static item => string.Equals(item, "Word x64", StringComparison.Ordinal)),
            mode + " installer handoff summary should list Word x64 as supported");
        AssertEqual(0, installerHandoffSummary.GetProperty("UnsupportedHosts").GetArrayLength(),
            mode + " installer handoff summary should not list unsupported hosts for the current machine");
        AssertTrue((installerHandoffSummary.GetProperty("UserFacingDecision").GetString() ?? string.Empty)
            .Contains("preview", StringComparison.OrdinalIgnoreCase),
            mode + " installer handoff summary should clearly state that execution remains preview-only");

        JsonElement liveEntrypointStatus = previewPlan.GetProperty("LiveEntrypointStatus");
        AssertEqual("PreviewOnly", liveEntrypointStatus.GetProperty("ExecutionMode").GetString(),
            mode + " preview plan should expose the current live-entrypoint status in preview-only mode");
        AssertEqual(true, liveEntrypointStatus.GetProperty("SharedCoreOwnsLiveRegistration").GetBoolean(),
            mode + " preview plan should confirm that the shared core now owns live registration for the current supported flow");
        AssertTrue(liveEntrypointStatus.GetProperty("CurrentEntrypoints")
            .EnumerateArray()
            .Select(static item => item.GetString())
            .Any(static item => string.Equals(item, "RegisterPlugin.ps1", StringComparison.Ordinal)),
            mode + " preview plan should document the current PowerShell live entrypoint");
        AssertTrue(liveEntrypointStatus.GetProperty("CurrentEntrypoints")
            .EnumerateArray()
            .Select(static item => item.GetString())
            .Any(static item => string.Equals(item, "RegisterPlugin.bat", StringComparison.Ordinal)),
            mode + " preview plan should document the current batch live entrypoint");
        AssertTrue(liveEntrypointStatus.GetProperty("ReroutedEntrypoints")
            .EnumerateArray()
            .Select(static item => item.GetString())
            .Any(static item => string.Equals(item, "Setup.iss", StringComparison.Ordinal)),
            mode + " preview plan should document that the installer entrypoint is rerouted");
        AssertEqual(0, liveEntrypointStatus.GetProperty("PendingEntrypoints").GetArrayLength(),
            mode + " preview plan should expose no pending live entrypoints for the current supported flow");
        AssertTrue((liveEntrypointStatus.GetProperty("MigrationDecision").GetString() ?? string.Empty)
            .Contains("delegate live registration", StringComparison.OrdinalIgnoreCase),
            mode + " preview plan should clearly state that live registration now delegates to the shared core");

        JsonElement migrationChecklist = previewPlan.GetProperty("MigrationChecklist");
        AssertEqual("PreviewOnly", migrationChecklist.GetProperty("ExecutionMode").GetString(),
            mode + " preview plan should expose a preview-only migration checklist");
        AssertEqual(true, migrationChecklist.GetProperty("ReadyToRerouteLiveEntrypoints").GetBoolean(),
            mode + " migration checklist should mark live-entrypoint reroute as complete for the current supported flow");
        AssertEqual(0, migrationChecklist.GetProperty("BlockingItems").GetArrayLength(),
            mode + " migration checklist should no longer report blocking items once the installer path is rerouted");
        AssertTrue(migrationChecklist.GetProperty("DeferredSupportTargets")
            .EnumerateArray()
            .Select(static item => item.GetString())
            .Any(static item => item?.Contains("WPS x64", StringComparison.Ordinal) == true),
            mode + " migration checklist should keep probe-pending hosts visible as deferred support targets");
        AssertTrue((migrationChecklist.GetProperty("OverallDecision").GetString() ?? string.Empty)
            .Contains("support expansion", StringComparison.OrdinalIgnoreCase),
            mode + " migration checklist should clearly distinguish reroute completion from future support expansion");

        JsonElement installerPreviewReport = previewPlan.GetProperty("InstallerPreviewReport");
        AssertEqual("PreviewOnly", installerPreviewReport.GetProperty("ExecutionMode").GetString(),
            mode + " preview plan should expose a preview-only installer preview report");
        AssertEqual(mode, installerPreviewReport.GetProperty("Operation").GetString(),
            mode + " installer preview report should preserve the requested operation");
        AssertEqual(1, installerPreviewReport.GetProperty("DetectedHostCount").GetInt32(),
            mode + " installer preview report should reflect the single detected host on the current machine");
        AssertEqual(1, installerPreviewReport.GetProperty("PreviewActionCount").GetInt32(),
            mode + " installer preview report should reflect the single preview action on the current machine");
        AssertEqual(true, installerPreviewReport.GetProperty("ReadyToRerouteLiveEntrypoints").GetBoolean(),
            mode + " installer preview report should expose that live reroute is complete for the current supported flow");
        AssertTrue((installerPreviewReport.GetProperty("SummaryText").GetString() ?? string.Empty)
            .Contains("Word x64", StringComparison.Ordinal),
            mode + " installer preview report should summarize the detected Word x64 host");

        JsonElement firstTarget = previewPlan.GetProperty("Targets")[0];
        string dispatchHandler = firstTarget.GetProperty("DispatchHandler").GetString() ?? string.Empty;
        JsonElement handlerPreview = firstTarget.GetProperty("HandlerPreview");
        AssertEqual("PreviewOnly", handlerPreview.GetProperty("ExecutionMode").GetString(),
            mode + " target should expose a preview-only low-level handler payload");
        JsonElement hostRuleSummary = handlerPreview.GetProperty("HostRuleSummary");
        AssertEqual("PreviewOnly", hostRuleSummary.GetProperty("ExecutionMode").GetString(),
            mode + " target should expose a preview-only host rule summary");
        AssertEqual("Word", hostRuleSummary.GetProperty("HostName").GetString(),
            mode + " host rule summary should preserve the detected host name");
        AssertEqual(mode, hostRuleSummary.GetProperty("Operation").GetString(),
            mode + " host rule summary should preserve the requested operation");
        AssertEqual("supported", hostRuleSummary.GetProperty("CurrentSupportStatus").GetString(),
            mode + " host rule summary should expose the current support state for the detected host");
        AssertEqual(true, hostRuleSummary.GetProperty("LiveExecutionAllowed").GetBoolean(),
            mode + " host rule summary should show that the current Word x64 host is eligible for live execution");
        AssertTrue(hostRuleSummary.GetProperty("EnablementCondition").GetString()?.Contains("supported", StringComparison.OrdinalIgnoreCase) == true,
            mode + " host rule summary should document the support gate for enablement");
        AssertTrue(hostRuleSummary.GetProperty("EnablementCondition").GetString()?.Contains("shared core can execute live", StringComparison.OrdinalIgnoreCase) == true,
            mode + " host rule summary should document that the shared core can now execute live Word x64 registration");
        AssertEqual(@"HKLM:\Software\Microsoft\Office\Word\Addins\WordTools.ThisAddIn",
            hostRuleSummary.GetProperty("RegistryTargets")[0].GetString(),
            mode + " host rule summary should expose the target Word add-in registry path");
        AssertEqual(hostRuleSummary.GetProperty("HandlerName").GetString(),
            topLevelHostRuleSummary.GetProperty("HandlerName").GetString(),
            mode + " grouped host rule summary should reuse the same handler identity as the handler-level summary");
        AssertEqual(hostRuleSummary.GetProperty("PreferredRegAsmInvoker").GetString(),
            topLevelHostRuleSummary.GetProperty("PreferredRegAsmInvoker").GetString(),
            mode + " grouped host rule summary should reuse the same RegAsm choice as the handler-level summary");
        AssertEqual(hostRuleSummary.GetProperty("RegistryTargets")[0].GetString(),
            topLevelHostRuleSummary.GetProperty("RegistryTargets")[0].GetString(),
            mode + " grouped host rule summary should reuse the same registry target as the handler-level summary");
        AssertEqual(hostRuleSummary.GetProperty("EnablementCondition").GetString(),
            topLevelHostRuleSummary.GetProperty("EnablementCondition").GetString(),
            mode + " grouped host rule summary should reuse the same enablement condition as the handler-level summary");
        if (string.Equals(mode, "Register", StringComparison.Ordinal))
        {
            AssertEqual("Register-WordHost", manifestItem.GetProperty("DispatchHandler").GetString(),
                "register operation manifest should point at the Word register handler");
            AssertEqual("Invoke-RegAsm64", manifestItem.GetProperty("RegAsmInvoker").GetString(),
                "register operation manifest should capture the 64-bit RegAsm invoker");
            AssertEqual("Register-WordHost", dispatchHandler,
                "current machine should route register preview through the Word-specific dispatch handler");
            AssertEqual("Register-WordHost", handlerPreview.GetProperty("HandlerName").GetString(),
                "current machine should expose the Word-specific register skeleton as the live handler target");
            AssertEqual("Invoke-RegAsm64", handlerPreview.GetProperty("RegAsmInvoker").GetString(),
                "current machine should resolve register preview through the 64-bit RegAsm helper");
            AssertEqual("Register-WordHost", hostRuleSummary.GetProperty("HandlerName").GetString(),
                "register host rule summary should point at the Word register handler");
            AssertEqual("Invoke-RegAsm64", hostRuleSummary.GetProperty("PreferredRegAsmInvoker").GetString(),
                "register host rule summary should capture the preferred 64-bit RegAsm path");
        }
        else
        {
            AssertEqual("Unregister-WordHost", manifestItem.GetProperty("DispatchHandler").GetString(),
                "unregister operation manifest should point at the Word unregister handler");
            AssertEqual("Invoke-RegAsm64", manifestItem.GetProperty("RegAsmInvoker").GetString(),
                "unregister operation manifest should capture the 64-bit RegAsm invoker");
            AssertEqual("Unregister-WordHost", dispatchHandler,
                "current machine should route unregister preview through the Word-specific dispatch handler");
            AssertEqual("Unregister-WordHost", handlerPreview.GetProperty("HandlerName").GetString(),
                "current machine should expose the Word-specific unregister skeleton as the live handler target");
            AssertEqual("Invoke-RegAsm64", handlerPreview.GetProperty("RegAsmInvoker").GetString(),
                "current machine should resolve unregister preview through the 64-bit RegAsm helper");
            AssertEqual("Unregister-WordHost", hostRuleSummary.GetProperty("HandlerName").GetString(),
                "unregister host rule summary should point at the Word unregister handler");
            AssertEqual("Invoke-RegAsm64", hostRuleSummary.GetProperty("PreferredRegAsmInvoker").GetString(),
                "unregister host rule summary should capture the preferred 64-bit RegAsm path");
        }
    }

    private static void TestProbeOutputStructureIncludesSupportReasonAndRegistrationView()
    {
        string core = ReadProjectSource("Installer.Core.ps1");

        AssertTrue(core.Contains("function Get-RegistrationView"),
            "probe core should define a registration-view resolver");
        AssertTrue(core.Contains("function Get-SupportDecision"),
            "probe core should define a support-decision resolver");
        AssertTrue(core.Contains("RegistrationView"),
            "probe output should expose the target registration view");
        AssertTrue(core.Contains("SupportReason"),
            "probe output should explain why a host is classified into its current support state");
    }

    private static void TestProbeOutputStructureIncludesValidationStageAndActivationRoute()
    {
        string core = ReadProjectSource("Installer.Core.ps1");

        AssertTrue(core.Contains("ValidationStage"),
            "probe output should expose the validation stage carried by the support matrix");
        AssertTrue(core.Contains("ActivationRoute"),
            "probe output should expose the activation route carried by the support matrix");

        string scriptPath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Installer.Core.ps1"));
        var startInfo = new System.Diagnostics.ProcessStartInfo
        {
            FileName = "powershell",
            Arguments = "-NoProfile -ExecutionPolicy Bypass -File \"" + scriptPath + "\" -Mode Probe -EvidenceLabel \"CurrentTurn-Probe-ValidationStage\"",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        using var process = System.Diagnostics.Process.Start(startInfo);
        string standardOutput = process.StandardOutput.ReadToEnd();
        string standardError = process.StandardError.ReadToEnd();
        process.WaitForExit();

        AssertEqual(0, process.ExitCode, "probe mode should complete successfully when surfacing validation stage and activation route. stderr: " + standardError);

        using JsonDocument probeResult = JsonDocument.Parse(standardOutput);
        JsonElement hosts = probeResult.RootElement.GetProperty("Hosts");

        AssertTrue(hosts.EnumerateArray().Any(entry =>
            string.Equals(entry.GetProperty("HostName").GetString(), "Word", StringComparison.Ordinal)
            && string.Equals(entry.GetProperty("ValidationStage").GetString(), "formal p0 passed", StringComparison.Ordinal)
            && string.Equals(entry.GetProperty("ActivationRoute").GetString(), "WordComAddin", StringComparison.Ordinal)),
            "detected Word hosts should expose the formal P0 stage and WordComAddin activation route on the current machine");

        bool detectedWpsHost = hosts.EnumerateArray().Any(entry =>
            string.Equals(entry.GetProperty("HostName").GetString(), "WPS", StringComparison.Ordinal));

        if (detectedWpsHost)
        {
            AssertTrue(hosts.EnumerateArray().Any(entry =>
                string.Equals(entry.GetProperty("HostName").GetString(), "WPS", StringComparison.Ordinal)
                && string.Equals(entry.GetProperty("ValidationStage").GetString(), "experimenting", StringComparison.Ordinal)
                && string.Equals(entry.GetProperty("ActivationRoute").GetString(), "WpsNativePlugin", StringComparison.Ordinal)),
                "detected WPS hosts should expose the experimenting validation stage and WpsNativePlugin activation route on the current machine");
        }
        else
        {
            JsonElement supportTargets = probeResult.RootElement
                .GetProperty("SupportState")
                .GetProperty("targets");

            AssertTrue(supportTargets.EnumerateArray().Any(entry =>
                string.Equals(entry.GetProperty("host").GetString(), "WPS", StringComparison.Ordinal)
                && string.Equals(entry.GetProperty("bitness").GetString(), "x86", StringComparison.Ordinal)
                && string.Equals(entry.GetProperty("validationStage").GetString(), "ui failed", StringComparison.Ordinal)
                && string.Equals(entry.GetProperty("activationRoute").GetString(), "WpsNativePlugin", StringComparison.Ordinal)),
                "probe output should preserve the WPS x86 failed-UI activation contract even when the current launch context does not detect a live WPS host");
        }
    }

    private static void TestProbeOutputStructureIncludesHostStateModelFields()
    {
        string core = ReadProjectSource("Installer.Core.ps1");

        AssertTrue(core.Contains("VersionLine"),
            "probe output should expose the resolved host version line");
        AssertTrue(core.Contains("InstallState"),
            "probe output should expose the host install-state classification");
        AssertTrue(core.Contains("UiEvidenceState"),
            "probe output should expose the host UI evidence state");
        AssertTrue(core.Contains("P0EvidenceState"),
            "probe output should expose the host P0 evidence state");
        AssertTrue(core.Contains("DiagnosticsBundleId"),
            "probe output should expose a host-scoped diagnostics bundle identifier");

        string scriptPath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Installer.Core.ps1"));
        var startInfo = new System.Diagnostics.ProcessStartInfo
        {
            FileName = "powershell",
            Arguments = "-NoProfile -ExecutionPolicy Bypass -File \"" + scriptPath + "\" -Mode Probe -EvidenceLabel \"CurrentTurn-Probe-HostStateModel\"",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        using var process = System.Diagnostics.Process.Start(startInfo);
        string standardOutput = process.StandardOutput.ReadToEnd();
        string standardError = process.StandardError.ReadToEnd();
        process.WaitForExit();

        AssertEqual(0, process.ExitCode, "probe mode should complete successfully when surfacing host state model fields. stderr: " + standardError);

        using JsonDocument probeResult = JsonDocument.Parse(standardOutput);
        JsonElement hosts = probeResult.RootElement.GetProperty("Hosts");

        JsonElement wordHost = hosts.EnumerateArray().First(entry =>
            string.Equals(entry.GetProperty("HostName").GetString(), "Word", StringComparison.Ordinal));

        AssertTrue(!string.IsNullOrWhiteSpace(wordHost.GetProperty("VersionLine").GetString()),
            "detected Word hosts should expose a non-empty version line");
        AssertEqual("detected", wordHost.GetProperty("InstallState").GetString(),
            "detected Word hosts should report a detected install state");
        AssertEqual("passed", wordHost.GetProperty("UiEvidenceState").GetString(),
            "the currently supported Word host should report passed UI evidence");
        AssertEqual("passed", wordHost.GetProperty("P0EvidenceState").GetString(),
            "the currently supported Word host should report passed P0 evidence");
        AssertTrue(wordHost.GetProperty("DiagnosticsBundleId").GetString()?.Contains("CurrentTurn-Probe-HostStateModel", StringComparison.Ordinal) == true,
            "detected Word hosts should stamp the current evidence label into the diagnostics bundle identifier");
    }

    private static void TestProbeOutputStructureIncludesWpsReconnaissanceDetails()
    {
        string core = ReadProjectSource("Installer.Core.ps1");

        AssertTrue(core.Contains("function Get-WpsReconData"),
            "probe core should define a dedicated WPS reconnaissance helper");
        AssertTrue(core.Contains("KnownPluginDirectories")
            && core.Contains("RegistryKeyPresence")
            && core.Contains("ConfigClues")
            && core.Contains("SetupPluginManifestSample")
            && core.Contains("DeclaredHosts")
            && core.Contains("PoolPackageRuntimeShape")
            && core.Contains("AddonStorageArchitectureSegment")
            && core.Contains("AddonStorageRoots")
            && core.Contains("AddonStorageCandidates")
            && core.Contains("PoolPackageHasRunInfoJson")
            && core.Contains("PoolPackageEntryDll")
            && core.Contains("PoolPackageEntryPoint")
            && core.Contains("PoolPackageLauncherType")
            && core.Contains("PoolPackageHasPluginProviderJson")
            && core.Contains("PoolPackageHasAttrPlg")
            && core.Contains("IndexSampleFileCount")
            && core.Contains("IndexPluginNameLiteralDetected")
            && core.Contains("IndexPluginNameUtf16LiteralDetected")
            && core.Contains("IndexPoolPackageRunInfoAppIdLiteralDetected")
            && core.Contains("IndexPoolPackageRunInfoAppIdUtf16LiteralDetected")
            && core.Contains("IndexPoolPackageEntryDllLiteralDetected")
            && core.Contains("IndexPoolPackageEntryDllUtf16LiteralDetected")
            && core.Contains("IndexPoolPackageEntryPointLiteralDetected")
            && core.Contains("IndexPoolPackageEntryPointUtf16LiteralDetected")
            && core.Contains("IndexPoolPackageLauncherTypeLiteralDetected")
            && core.Contains("IndexPoolPackageLauncherTypeUtf16LiteralDetected")
            && core.Contains("SetupPluginAuthInfoBinaryWrapped")
            && core.Contains("ConfiguredComAddinsDialogClass")
            && core.Contains("ConfiguredComAddinsDialogHosts")
            && core.Contains("ComAddinsDialogHostModule")
            && core.Contains("ApplicationApiSurface")
            && core.Contains("ComAddinsCommandDatabaseEvidence")
            && core.Contains("ComAddinsCommandDatabaseSamples")
            && core.Contains("LikelyExternalAddinRegistryRoot")
            && core.Contains("ExistingExternalAddinEntryCount")
            && core.Contains("ExistingExternalAddinEntrySample")
            && core.Contains("NonEmptyExternalAddinEntryCount")
            && core.Contains("NonEmptyExternalAddinEntries")
            && core.Contains("ExternalAddinEntryResolutionSamples")
            && core.Contains("ResolvedExternalAddinProgIdSampleCount")
            && core.Contains("ResolvedExternalAddinProgIdTotalCount")
            && core.Contains("SuspiciousModuleFiles")
            && core.Contains("SuspiciousModuleSamples")
            && core.Contains("NamedExportSample")
            && core.Contains("IndexedPluginStores")
            && core.Contains("ListV3ShardDirectorySamples")
            && core.Contains("NumericShardDirectoryCount")
            && core.Contains("SampleShardIdsLength")
            && core.Contains("SampleShardDataLength")
            && core.Contains("SampleShardDirectoryNameLiteralDetectedInPluginList")
            && core.Contains("SampleShardDirectoryNameUtf16LiteralDetectedInPluginList")
            && core.Contains("SampleShardDirectoryNameUInt32LittleEndianDetectedInPluginList")
            && core.Contains("SampleShardIdLiteralDetectedInPoolMetadata")
            && core.Contains("NativePluginPackageMetadataFiles")
            && core.Contains("NativePluginMetadataSignals")
            && core.Contains("NativeDllPackageSamples")
            && core.Contains("VbaRuntimeInstallArtifacts")
            && core.Contains("InternalComTypeLibEvidence"),
            "WPS reconnaissance should capture plugin directories, setupplugin clues, registry-key presence, COM add-ins dialog evidence, application API surface clues, external add-in root evidence, suspicious module files, plugin indexes, package-metadata signals, VBA runtime artifacts, command-database evidence, and internal COM typelib traces");

        string scriptPath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Installer.Core.ps1"));
        var startInfo = new System.Diagnostics.ProcessStartInfo
        {
            FileName = "powershell",
            Arguments = "-NoProfile -ExecutionPolicy Bypass -File \"" + scriptPath + "\" -Mode Probe -EvidenceLabel \"CurrentMachine-Word64-WpsX86-20260520\"",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        using var process = System.Diagnostics.Process.Start(startInfo);
        string standardOutput = process.StandardOutput.ReadToEnd();
        string standardError = process.StandardError.ReadToEnd();
        process.WaitForExit();

        AssertEqual(0, process.ExitCode, "probe mode should complete successfully when gathering WPS reconnaissance details. stderr: " + standardError);

        using JsonDocument probeResult = JsonDocument.Parse(standardOutput);
        JsonElement hosts = probeResult.RootElement.GetProperty("Hosts");
        foreach (JsonElement host in hosts.EnumerateArray())
        {
            if (!string.Equals(host.GetProperty("HostName").GetString(), "WPS", StringComparison.Ordinal))
            {
                continue;
            }

            AssertTrue(host.TryGetProperty("WpsRecon", out JsonElement recon) && recon.ValueKind == JsonValueKind.Object,
                "detected WPS hosts should include a WPS reconnaissance object");
            AssertTrue(recon.TryGetProperty("KnownPluginDirectories", out JsonElement pluginDirs)
                && pluginDirs.ValueKind == JsonValueKind.Array
                && pluginDirs.GetArrayLength() > 0,
                "WPS reconnaissance should list discovered plugin framework directories");
            AssertTrue(recon.TryGetProperty("AddonStorageArchitectureSegment", out JsonElement addonStorageArchitectureSegment)
                && string.Equals(addonStorageArchitectureSegment.GetString(), "win-i386", StringComparison.Ordinal),
                "current-machine WPS x86 reconnaissance should expose that addon storage selection is currently routed through the win-i386 architecture segment");
            AssertTrue(recon.TryGetProperty("AddonStorageRoots", out JsonElement addonStorageRoots)
                && addonStorageRoots.ValueKind == JsonValueKind.Object
                && (addonStorageRoots.GetProperty("PoolRoot").GetString() ?? string.Empty).EndsWith(@"\pool\win-i386", StringComparison.OrdinalIgnoreCase)
                && (addonStorageRoots.GetProperty("ListRoot").GetString() ?? string.Empty).EndsWith(@"\list\win-i386", StringComparison.OrdinalIgnoreCase)
                && (addonStorageRoots.GetProperty("ListV3Root").GetString() ?? string.Empty).EndsWith(@"\listV3\win-i386", StringComparison.OrdinalIgnoreCase),
                "current-machine WPS x86 reconnaissance should expose the resolved addon pool/list/listV3 roots that match the selected architecture segment");
            AssertTrue(recon.TryGetProperty("AddonStorageCandidates", out JsonElement addonStorageCandidates)
                && addonStorageCandidates.ValueKind == JsonValueKind.Array
                && addonStorageCandidates.EnumerateArray().Any(entry =>
                    string.Equals(entry.GetProperty("ArchitectureSegment").GetString(), "win-i386", StringComparison.Ordinal)
                    && entry.GetProperty("PoolRootPresent").ValueKind == JsonValueKind.True
                    && entry.GetProperty("ListRootPresent").ValueKind == JsonValueKind.True
                    && entry.GetProperty("ListV3RootPresent").ValueKind == JsonValueKind.True)
                && addonStorageCandidates.EnumerateArray().Any(entry =>
                    string.Equals(entry.GetProperty("ArchitectureSegment").GetString(), "win-x64", StringComparison.Ordinal)
                    && entry.GetProperty("PoolRootPresent").ValueKind == JsonValueKind.False
                    && entry.GetProperty("ListRootPresent").ValueKind == JsonValueKind.False
                    && entry.GetProperty("ListV3RootPresent").ValueKind == JsonValueKind.False),
                "current-machine WPS x86 reconnaissance should expose both the present win-i386 addon roots and the currently absent win-x64 sibling roots");
            AssertTrue(recon.TryGetProperty("RegistryKeyPresence", out JsonElement registryPresence)
                && registryPresence.ValueKind == JsonValueKind.Object,
                "WPS reconnaissance should expose the observed WPS registry-key states");
            AssertTrue(recon.TryGetProperty("SetupPluginManifestSample", out JsonElement setupPluginManifestSample)
                && setupPluginManifestSample.ValueKind == JsonValueKind.Array
                && setupPluginManifestSample.GetArrayLength() > 0,
                "WPS reconnaissance should expose sampled setupplugin manifest entries");
            AssertTrue(setupPluginManifestSample.EnumerateArray().Any(entry =>
                string.Equals(entry.GetProperty("Name").GetString(), "kdocerjsapi20", StringComparison.Ordinal)
                && (entry.GetProperty("Host").GetString() ?? string.Empty).Contains("wps", StringComparison.OrdinalIgnoreCase)
                && entry.GetProperty("DeclaredHosts").EnumerateArray().Any(name => string.Equals(name.GetString(), "wps", StringComparison.OrdinalIgnoreCase))
                && entry.GetProperty("DeclaredHosts").EnumerateArray().Any(name => string.Equals(name.GetString(), "wpp", StringComparison.OrdinalIgnoreCase))
                && entry.GetProperty("DeclaredHosts").EnumerateArray().Any(name => string.Equals(name.GetString(), "et", StringComparison.OrdinalIgnoreCase))
                && entry.GetProperty("DeclaredHosts").EnumerateArray().Any(name => string.Equals(name.GetString(), "pdf", StringComparison.OrdinalIgnoreCase))
                && string.Equals(entry.GetProperty("Type").GetString(), "dll", StringComparison.Ordinal)
                && (entry.GetProperty("PoolPackageDirectory").GetString() ?? string.Empty).Contains("kdocerjsapi20_", StringComparison.OrdinalIgnoreCase)
                && string.Equals(entry.GetProperty("PoolPackageRuntimeShape").GetString(), "RunInfoDrivenWebOrJsApi", StringComparison.Ordinal)
                && entry.GetProperty("PoolPackageHasRunInfoJson").ValueKind == JsonValueKind.True
                && string.Equals(entry.GetProperty("PoolPackageRunInfoAppId").GetString(), "picture_resourceshop_split", StringComparison.Ordinal)
                && string.Equals(entry.GetProperty("PoolPackageEntryDll").GetString(), "kdocerjsapi20.dll", StringComparison.Ordinal)
                && string.Equals(entry.GetProperty("PoolPackageEntryPoint").GetString(), "CreateSplitAppWidget", StringComparison.Ordinal)
                && string.Equals(entry.GetProperty("PoolPackageLauncherType").GetString(), "proxyFrame", StringComparison.Ordinal)
                && entry.GetProperty("IndexSampleFileCount").GetInt32() > 0
                && entry.GetProperty("IndexPluginNameLiteralDetected").ValueKind == JsonValueKind.False
                && entry.GetProperty("IndexPluginNameUtf16LiteralDetected").ValueKind == JsonValueKind.False
                && entry.GetProperty("IndexPoolPackageRunInfoAppIdLiteralDetected").ValueKind == JsonValueKind.False
                && entry.GetProperty("IndexPoolPackageRunInfoAppIdUtf16LiteralDetected").ValueKind == JsonValueKind.False
                && entry.GetProperty("IndexPoolPackageEntryDllLiteralDetected").ValueKind == JsonValueKind.False
                && entry.GetProperty("IndexPoolPackageEntryDllUtf16LiteralDetected").ValueKind == JsonValueKind.False
                && entry.GetProperty("IndexPoolPackageEntryPointLiteralDetected").ValueKind == JsonValueKind.False
                && entry.GetProperty("IndexPoolPackageEntryPointUtf16LiteralDetected").ValueKind == JsonValueKind.False
                && entry.GetProperty("IndexPoolPackageLauncherTypeLiteralDetected").ValueKind == JsonValueKind.False
                && entry.GetProperty("IndexPoolPackageLauncherTypeUtf16LiteralDetected").ValueKind == JsonValueKind.False
                && entry.TryGetProperty("PoolPackageHasPluginProviderJson", out _)),
                "WPS reconnaissance should expose sampled setupplugin declared hosts together with a runinfo-driven runtime shape, mapped-package app identity and entry metadata, and binary-index ASCII plus UTF-16LE literal-search results for multi-host web-or-JSAPI component plugins");
            AssertTrue(setupPluginManifestSample.EnumerateArray().Any(entry =>
                string.Equals(entry.GetProperty("Name").GetString(), "kwpsaiwordtool", StringComparison.Ordinal)
                && entry.GetProperty("PoolPackageDirectoryPresent").ValueKind == JsonValueKind.True
                && entry.GetProperty("DeclaredHosts").GetArrayLength() == 1
                && string.Equals(entry.GetProperty("DeclaredHosts")[0].GetString(), "wps", StringComparison.OrdinalIgnoreCase)
                && string.Equals(entry.GetProperty("PoolPackageRuntimeShape").GetString(), "DllAttrNativeModule", StringComparison.Ordinal)
                && entry.GetProperty("PoolPackageHasRunInfoJson").ValueKind == JsonValueKind.False
                && entry.GetProperty("PoolPackageHasRunIni").ValueKind == JsonValueKind.False
                && entry.GetProperty("IndexSampleFileCount").GetInt32() > 0
                && entry.GetProperty("IndexPluginNameLiteralDetected").ValueKind == JsonValueKind.False
                && entry.GetProperty("IndexPluginNameUtf16LiteralDetected").ValueKind == JsonValueKind.False
                && entry.GetProperty("PoolPackageHasAttrPlg").ValueKind == JsonValueKind.True),
                "WPS reconnaissance should surface when a setupplugin declaration resolves to a WPS-only DLL-plus-attr package, distinguish that runtime shape from runinfo-driven web packages, and show that sampled binary indexes do not expose the plugin name as a plain ASCII or UTF-16LE literal");
            AssertTrue(recon.TryGetProperty("SetupPluginAuthInfoBinaryWrapped", out JsonElement setupPluginAuthInfoBinaryWrapped)
                && setupPluginAuthInfoBinaryWrapped.ValueKind == JsonValueKind.True,
                "WPS reconnaissance should identify setuppluginauthinfo.json as a binary-wrapped payload rather than plain JSON");
            AssertTrue(recon.TryGetProperty("ConfiguredComAddinsDialogClass", out JsonElement dialogClass)
                && string.Equals(dialogClass.GetString(), "TCOMAddinsDlg.UnicodeClass", StringComparison.Ordinal),
                "WPS reconnaissance should surface the shared COM add-ins dialog class configured in WPS, WPP, and ET");
            AssertTrue(recon.TryGetProperty("ConfiguredComAddinsDialogHosts", out JsonElement dialogHosts)
                && dialogHosts.ValueKind == JsonValueKind.Array
                && dialogHosts.EnumerateArray().Any(entry => string.Equals(entry.GetString(), "WPS", StringComparison.Ordinal))
                && dialogHosts.EnumerateArray().Any(entry => string.Equals(entry.GetString(), "WPP", StringComparison.Ordinal))
                && dialogHosts.EnumerateArray().Any(entry => string.Equals(entry.GetString(), "ET", StringComparison.Ordinal)),
                "WPS reconnaissance should identify that WPS, WPP, and ET share the COM add-ins dialog mapping");
            AssertTrue(recon.TryGetProperty("ComAddinsDialogHostModule", out JsonElement dialogHostModule)
                && (dialogHostModule.GetString() ?? string.Empty).EndsWith("kshell.dll", StringComparison.OrdinalIgnoreCase),
                "WPS reconnaissance should identify kshell.dll as the shared COM add-ins dialog host candidate");
            AssertTrue(recon.TryGetProperty("ApplicationApiSurface", out JsonElement apiSurface)
                && apiSurface.ValueKind == JsonValueKind.Object,
                "WPS reconnaissance should surface the sampled application API exposure for AddIns and COMAddIns");
            AssertTrue(apiSurface.TryGetProperty("WppHasComAddIns", out JsonElement wppHasComAddIns)
                && wppHasComAddIns.ValueKind == JsonValueKind.True,
                "WPS reconnaissance should show that the sampled WPP application API explicitly exposes COMAddIns");
            AssertTrue(apiSurface.TryGetProperty("WpsHasAddIns", out JsonElement wpsHasAddIns)
                && wpsHasAddIns.ValueKind == JsonValueKind.True,
                "WPS reconnaissance should show that the sampled WPS application API exposes AddIns");
            AssertTrue(apiSurface.TryGetProperty("WpsHasComAddIns", out JsonElement wpsHasComAddIns)
                && wpsHasComAddIns.ValueKind == JsonValueKind.True,
                "WPS reconnaissance should show that the sampled WPS application API explicitly exposes COMAddIns");
            AssertTrue(apiSurface.TryGetProperty("EtHasComAddIns", out JsonElement etHasComAddIns)
                && etHasComAddIns.ValueKind == JsonValueKind.True,
                "WPS reconnaissance should show that the sampled ET application API explicitly exposes COMAddIns");
            AssertTrue(recon.TryGetProperty("LikelyExternalAddinRegistryRoot", out JsonElement externalRoot)
                && string.Equals(externalRoot.GetString(), @"HKCU:\Software\Kingsoft\Office\WPS\AddinsWl", StringComparison.Ordinal),
                "WPS reconnaissance should surface the strongest candidate external add-in registry root");
            AssertTrue(recon.TryGetProperty("ExistingExternalAddinEntryCount", out JsonElement externalEntryCount)
                && externalEntryCount.GetInt32() > 0,
                "WPS reconnaissance should count existing add-in entries when the candidate root is populated");
            AssertTrue(recon.TryGetProperty("ExistingExternalAddinEntrySample", out JsonElement externalEntrySample)
                && externalEntrySample.ValueKind == JsonValueKind.Array
                && externalEntrySample.GetArrayLength() > 0,
                "WPS reconnaissance should surface sample existing add-in entry names from the candidate root");
            AssertTrue(recon.TryGetProperty("NonEmptyExternalAddinEntryCount", out JsonElement nonEmptyEntryCount)
                && nonEmptyEntryCount.GetInt32() == 2,
                "WPS reconnaissance should count the AddinsWl entries that carry non-empty payload values");
            AssertTrue(recon.TryGetProperty("NonEmptyExternalAddinEntries", out JsonElement nonEmptyEntries)
                && nonEmptyEntries.ValueKind == JsonValueKind.Array
                && nonEmptyEntries.GetArrayLength() == 2
                && nonEmptyEntries.EnumerateArray().All(entry => entry.TryGetProperty("Name", out _) && entry.TryGetProperty("Value", out _)),
                "WPS reconnaissance should surface the non-empty AddinsWl entries as name/value evidence");
            AssertTrue(externalEntrySample.EnumerateArray().Any(entry =>
                (entry.GetString() ?? string.Empty).Contains("Word", StringComparison.OrdinalIgnoreCase)
                || (entry.GetString() ?? string.Empty).Contains("WPS", StringComparison.OrdinalIgnoreCase)
                || (entry.GetString() ?? string.Empty).Contains("HyWordAI", StringComparison.OrdinalIgnoreCase)),
                "WPS reconnaissance should show ProgID-like samples from the candidate external add-in root");
            AssertTrue(recon.TryGetProperty("ExternalAddinEntryResolutionSamples", out JsonElement resolutionSamples)
                && resolutionSamples.ValueKind == JsonValueKind.Array
                && resolutionSamples.GetArrayLength() > 0,
                "WPS reconnaissance should expose sampled resolution checks for candidate external add-in entries");
            AssertTrue(resolutionSamples.EnumerateArray().All(entry =>
                entry.TryGetProperty("Name", out _)
                && entry.TryGetProperty("HkcrProgIdPresent", out _)
                && entry.TryGetProperty("HkcuClassesProgIdPresent", out _)
                && entry.TryGetProperty("HklmClassesProgIdPresent", out _)),
                "WPS reconnaissance resolution samples should expose the key ProgID presence flags");
            AssertTrue(recon.TryGetProperty("ResolvedExternalAddinProgIdSampleCount", out JsonElement resolvedSampleCount)
                && resolvedSampleCount.ValueKind == JsonValueKind.Number,
                "WPS reconnaissance should count how many sampled external add-in entry names currently resolve to live ProgID keys");
            AssertTrue(recon.TryGetProperty("ResolvedExternalAddinProgIdTotalCount", out JsonElement resolvedTotalCount)
                && resolvedTotalCount.GetInt32() == 0,
                "WPS reconnaissance should surface that none of the current-machine AddinsWl entry names resolve to live ProgID roots");
            AssertTrue(recon.TryGetProperty("SuspiciousModuleFiles", out JsonElement suspiciousModules)
                && suspiciousModules.ValueKind == JsonValueKind.Array
                && suspiciousModules.GetArrayLength() > 0,
                "WPS reconnaissance should surface suspicious addon bridge module files");
            AssertTrue(suspiciousModules.EnumerateArray().Any(entry =>
                (entry.GetString() ?? string.Empty).Contains("kwpsremixsdksrv", StringComparison.OrdinalIgnoreCase)
                || (entry.GetString() ?? string.Empty).Contains("kvbarunner", StringComparison.OrdinalIgnoreCase)
                || (entry.GetString() ?? string.Empty).Contains("kpluginmanager", StringComparison.OrdinalIgnoreCase)),
                "WPS reconnaissance should surface at least one of the observed bridge modules");
            AssertTrue(recon.TryGetProperty("SuspiciousModuleSamples", out JsonElement suspiciousModuleSamples)
                && suspiciousModuleSamples.ValueKind == JsonValueKind.Array
                && suspiciousModuleSamples.GetArrayLength() > 0
                && suspiciousModuleSamples.EnumerateArray().Any(entry =>
                    (entry.GetProperty("Path").GetString() ?? string.Empty).Contains("kpluginmanager.dll", StringComparison.OrdinalIgnoreCase)
                    && entry.GetProperty("NamedExportCount").GetInt32() > 0
                    && entry.GetProperty("NamedExportSample").EnumerateArray().Any(name =>
                        string.Equals(name.GetString(), "parsePluginListItem", StringComparison.Ordinal)
                        || string.Equals(name.GetString(), "KPluginManager", StringComparison.Ordinal))
                    && entry.GetProperty("HasDllRegisterServer").ValueKind == JsonValueKind.False
                    && entry.GetProperty("HasClsidLiteral").ValueKind == JsonValueKind.False),
                "WPS reconnaissance should structure kpluginmanager-style bridge-module samples with parser-like exports while showing they do not currently look like classic COM add-ins");
            AssertTrue(recon.TryGetProperty("IndexedPluginStores", out JsonElement indexedStores)
                && indexedStores.ValueKind == JsonValueKind.Object
                && indexedStores.TryGetProperty("ListV3RootPresent", out JsonElement listV3Present)
                && listV3Present.ValueKind == JsonValueKind.True,
                "WPS reconnaissance should expose the versioned addon index store");
            AssertTrue(indexedStores.TryGetProperty("PluginListFiles", out JsonElement pluginLists)
                && pluginLists.ValueKind == JsonValueKind.Array
                && pluginLists.GetArrayLength() > 0,
                "WPS reconnaissance should enumerate discovered pluginlist manifests");
            AssertTrue(indexedStores.TryGetProperty("ListV3ShardDirectorySamples", out JsonElement shardDirectorySamples)
                && shardDirectorySamples.ValueKind == JsonValueKind.Array
                && shardDirectorySamples.GetArrayLength() > 0
                && shardDirectorySamples.EnumerateArray().Any(entry =>
                    entry.GetProperty("NumericShardDirectoryCount").GetInt32() > 0
                    && !string.IsNullOrWhiteSpace(entry.GetProperty("SampleShardDirectoryName").GetString())
                    && entry.GetProperty("SampleShardIdsLength").GetInt64() > 0
                    && entry.GetProperty("SampleShardDataLength").GetInt64() > 0
                    && entry.TryGetProperty("SampleShardDirectoryNameLiteralDetectedInPluginList", out JsonElement sampleShardDirectoryNameLiteralDetectedInPluginList)
                    && sampleShardDirectoryNameLiteralDetectedInPluginList.ValueKind == JsonValueKind.False
                    && entry.TryGetProperty("SampleShardDirectoryNameUtf16LiteralDetectedInPluginList", out JsonElement sampleShardDirectoryNameUtf16LiteralDetectedInPluginList)
                    && sampleShardDirectoryNameUtf16LiteralDetectedInPluginList.ValueKind == JsonValueKind.False
                    && entry.TryGetProperty("SampleShardDirectoryNameUInt32LittleEndianDetectedInPluginList", out JsonElement sampleShardDirectoryNameUInt32LittleEndianDetectedInPluginList)
                    && sampleShardDirectoryNameUInt32LittleEndianDetectedInPluginList.ValueKind == JsonValueKind.False
                    && entry.TryGetProperty("SampleShardIdLiteralDetectedInPoolMetadata", out JsonElement sampleShardIdLiteralDetectedInPoolMetadata)
                    && sampleShardIdLiteralDetectedInPoolMetadata.ValueKind == JsonValueKind.False),
                "WPS reconnaissance should expose versioned listV3 shard-layout samples, including shard counts, sampled ids/data file sizes, and whether a sampled shard ID is directly visible in pluginlist as text, as a UInt32 little-endian pattern, or in pool metadata");
            AssertTrue(indexedStores.TryGetProperty("BinaryHeaderSamples", out JsonElement binaryHeaderSamples)
                && binaryHeaderSamples.ValueKind == JsonValueKind.Array
                && binaryHeaderSamples.GetArrayLength() > 0
                && binaryHeaderSamples.EnumerateArray().Any(entry =>
                    (entry.GetProperty("HeaderAsciiPrefix").GetString() ?? string.Empty).StartsWith("kplugin", StringComparison.Ordinal)),
                "WPS reconnaissance should show that sampled pluginlist and shard files carry a shared kplugin header");
            AssertTrue(recon.TryGetProperty("NativePluginPackageMetadataFiles", out JsonElement nativePluginMetadata)
                && nativePluginMetadata.ValueKind == JsonValueKind.Array
                && nativePluginMetadata.GetArrayLength() > 0,
                "WPS reconnaissance should surface the native addon package metadata files observed under the WPS pool store");
            AssertTrue(nativePluginMetadata.EnumerateArray().Any(entry =>
                (entry.GetString() ?? string.Empty).Contains("plugin-provider.json", StringComparison.OrdinalIgnoreCase)
                || (entry.GetString() ?? string.Empty).Contains("runinfo.json", StringComparison.OrdinalIgnoreCase)
                || (entry.GetString() ?? string.Empty).Contains("config.json", StringComparison.OrdinalIgnoreCase)
                || (entry.GetString() ?? string.Empty).Contains("run.ini", StringComparison.OrdinalIgnoreCase)),
                "WPS reconnaissance should show at least one native addon package metadata file that is stronger than a pure AddinsWl-only story");
            AssertTrue(recon.TryGetProperty("NativePluginMetadataSignals", out JsonElement metadataSignals)
                && metadataSignals.ValueKind == JsonValueKind.Object,
                "WPS reconnaissance should expose structured signals extracted from sampled native package metadata");
            AssertTrue(metadataSignals.TryGetProperty("RunInfoManifestSamples", out JsonElement runInfoSamples)
                && runInfoSamples.ValueKind == JsonValueKind.Array
                && runInfoSamples.EnumerateArray().Any(entry =>
                    string.Equals(entry.GetProperty("EntryDll").GetString(), "kdocerjsapi20.dll", StringComparison.Ordinal)
                    && string.Equals(entry.GetProperty("EntryPoint").GetString(), "CreateSplitAppWidget", StringComparison.Ordinal)
                    && string.Equals(entry.GetProperty("LauncherType").GetString(), "proxyFrame", StringComparison.Ordinal)),
                "WPS reconnaissance should extract entryDll, entryPoint, and launcherType from sampled runinfo manifests");
            AssertTrue(metadataSignals.TryGetProperty("ConfigManifestSamples", out JsonElement configSamples)
                && configSamples.ValueKind == JsonValueKind.Array
                && configSamples.EnumerateArray().Any(entry =>
                    string.Equals(entry.GetProperty("OfficeType").GetString(), "w", StringComparison.Ordinal)),
                "WPS reconnaissance should extract office_type routing from sampled config manifests");
            AssertTrue(metadataSignals.TryGetProperty("RunIniManifestSamples", out JsonElement runIniSamples)
                && runIniSamples.ValueKind == JsonValueKind.Array
                && runIniSamples.EnumerateArray().Any(entry =>
                    string.Equals(entry.GetProperty("Entry").GetString(), "%workingroot%/index.html", StringComparison.Ordinal)
                    && string.Equals(entry.GetProperty("IsLoadOnline").GetString(), "0", StringComparison.Ordinal)),
                "WPS reconnaissance should extract run.ini entry and load flags from sampled manifests");
            AssertTrue(metadataSignals.TryGetProperty("PreloadFileSamples", out JsonElement preloadFileSamples)
                && preloadFileSamples.ValueKind == JsonValueKind.Array
                && preloadFileSamples.EnumerateArray().Any(entry =>
                    entry.GetProperty("ContainsWriterBundles").ValueKind == JsonValueKind.True
                    && entry.GetProperty("PcBundleCount").GetInt32() > 0),
                "WPS reconnaissance should detect writer-specific preload bundles in sampled native web packages");
            AssertTrue(metadataSignals.TryGetProperty("BinaryWrappedMetadataSamples", out JsonElement binaryWrappedMetadataSamples)
                && binaryWrappedMetadataSamples.ValueKind == JsonValueKind.Array
                && binaryWrappedMetadataSamples.EnumerateArray().Any(entry =>
                    (entry.GetProperty("Path").GetString() ?? string.Empty).Contains("plugin-provider.json", StringComparison.OrdinalIgnoreCase)
                    || (entry.GetProperty("Path").GetString() ?? string.Empty).Contains("__attr.plg", StringComparison.OrdinalIgnoreCase)),
                "WPS reconnaissance should distinguish binary-wrapped metadata blobs from plain JSON/INI manifests");
            AssertTrue(metadataSignals.TryGetProperty("NativeDllPackageSamples", out JsonElement nativeDllPackageSamples)
                && nativeDllPackageSamples.ValueKind == JsonValueKind.Array
                && nativeDllPackageSamples.GetArrayLength() > 0,
                "WPS reconnaissance should expose sampled native DLL-style addon packages alongside web-style packages");
            AssertTrue(nativeDllPackageSamples.EnumerateArray().Any(entry =>
                (entry.GetProperty("DllPath").GetString() ?? string.Empty).Contains("kdocerjsapi20.dll", StringComparison.OrdinalIgnoreCase)
                && entry.GetProperty("NamedExportCount").GetInt32() > 0
                && entry.GetProperty("NamedExportSample").EnumerateArray().Any(name => string.Equals(name.GetString(), "GetExtensionJsApiObj", StringComparison.Ordinal))
                && entry.GetProperty("HasDllRegisterServer").ValueKind == JsonValueKind.False
                && entry.GetProperty("HasProgIdLiteral").ValueKind == JsonValueKind.False),
                "WPS reconnaissance should show that the sampled kdocerjsapi20 package uses custom JS/API-style exports rather than COM-registration markers");
            AssertTrue(nativeDllPackageSamples.EnumerateArray().Any(entry =>
                (entry.GetProperty("DllPath").GetString() ?? string.Empty).Contains("kwpsaiwordtool.dll", StringComparison.OrdinalIgnoreCase)
                && entry.GetProperty("HasRunInfoJson").ValueKind == JsonValueKind.False
                && entry.GetProperty("HasConfigJson").ValueKind == JsonValueKind.False
                && entry.GetProperty("HasRunIni").ValueKind == JsonValueKind.False
                && entry.GetProperty("HasAttrPlg").ValueKind == JsonValueKind.True
                && entry.GetProperty("NamedExportCount").GetInt32() == 0
                && entry.GetProperty("HasDllGetClassObject").ValueKind == JsonValueKind.False
                && entry.GetProperty("HasClsidLiteral").ValueKind == JsonValueKind.False),
                "WPS reconnaissance should show that the sampled kwpsaiwordtool package is a DLL-plus-attr package without the sampled web manifests or classic COM markers");
            AssertTrue(recon.TryGetProperty("VbaRuntimeInstallArtifacts", out JsonElement vbaArtifacts)
                && vbaArtifacts.ValueKind == JsonValueKind.Array
                && vbaArtifacts.EnumerateArray().Any(entry =>
                    (entry.GetString() ?? string.Empty).Contains("vba7.zip", StringComparison.OrdinalIgnoreCase)),
                "WPS reconnaissance should expose the bundled VBA runtime installer artifacts");
            AssertTrue(recon.TryGetProperty("InternalComTypeLibEvidence", out JsonElement typeLibEvidence)
                && typeLibEvidence.ValueKind == JsonValueKind.Array
                && typeLibEvidence.GetArrayLength() > 0,
                "WPS reconnaissance should expose internal COM typelib traces when present");
            AssertTrue(recon.TryGetProperty("ComAddinsCommandDatabaseEvidence", out JsonElement commandDatabaseEvidence)
                && commandDatabaseEvidence.ValueKind == JsonValueKind.Array
                && commandDatabaseEvidence.GetArrayLength() > 0,
                "WPS reconnaissance should surface the command-database files that still contain COMAddIns UI evidence");
            AssertTrue(recon.TryGetProperty("ComAddinsCommandDatabaseSamples", out JsonElement commandDatabaseSamples)
                && commandDatabaseSamples.ValueKind == JsonValueKind.Array
                && commandDatabaseSamples.GetArrayLength() > 0,
                "WPS reconnaissance should surface sampled command-database details including package version, host database name, and literal-search hits");
            AssertTrue(commandDatabaseSamples.EnumerateArray().Any(sample =>
                    sample.TryGetProperty("PackageVersion", out _)
                    && sample.TryGetProperty("HostDatabaseName", out _)
                    && sample.TryGetProperty("ContainsComAddIns", out JsonElement contains) && contains.GetBoolean()
                    && sample.TryGetProperty("ContainsKdocerjsapi20", out _)
                    && sample.TryGetProperty("ContainsKwpsaiwordtool", out _)
                    && sample.TryGetProperty("ContainsPictureResourceshopSplit", out _)),
                "WPS reconnaissance command-database samples should expose PackageVersion, HostDatabaseName, ContainsComAddIns=true, and literal-search results for known plugin names");
            AssertTrue(commandDatabaseSamples.EnumerateArray().Any(sample =>
                    sample.TryGetProperty("ContainsKdocerjsapi20", out JsonElement kdoc) && kdoc.GetBoolean()
                    || sample.TryGetProperty("ContainsKwpsaiwordtool", out JsonElement kword) && kword.GetBoolean()),
                "WPS reconnaissance should surface at least one command database that contains kdocerjsapi20 or kwpsaiwordtool as a plain literal");
            AssertTrue(commandDatabaseSamples.EnumerateArray().All(sample =>
                    !(sample.TryGetProperty("ContainsKdocerjsapi20EntryDll", out JsonElement entryDll) && entryDll.GetBoolean())
                    && !(sample.TryGetProperty("ContainsKdocerjsapi20EntryPoint", out JsonElement entryPoint) && entryPoint.GetBoolean())
                    && !(sample.TryGetProperty("ContainsKdocerjsapi20LauncherType", out JsonElement launcherType) && launcherType.GetBoolean())),
                "WPS reconnaissance command-database samples should confirm that kdocerjsapi20.dll, CreateSplitAppWidget, and proxyFrame are NOT present as plain literals — entry DLL, entry point, and launcher type tokens are absent from sampled command databases");
            AssertTrue(recon.TryGetProperty("Conclusion", out JsonElement conclusion)
                && (conclusion.GetString() ?? string.Empty).Contains("HKCU:\\Software\\Kingsoft\\Office\\WPS\\AddinsWl", StringComparison.OrdinalIgnoreCase)
                && (conclusion.GetString() ?? string.Empty).Contains("kshell.dll", StringComparison.OrdinalIgnoreCase)
                && (conclusion.GetString() ?? string.Empty).Contains("TCOMAddinsDlg.UnicodeClass", StringComparison.OrdinalIgnoreCase)
                && (conclusion.GetString() ?? string.Empty).Contains("WPS, WPP, and ET", StringComparison.OrdinalIgnoreCase)
                && (conclusion.GetString() ?? string.Empty).Contains("Application.COMAddIns", StringComparison.OrdinalIgnoreCase)
                && (conclusion.GetString() ?? string.Empty).Contains("zero resolved to live ProgID roots", StringComparison.OrdinalIgnoreCase)
                && (conclusion.GetString() ?? string.Empty).Contains("version-gate-like", StringComparison.OrdinalIgnoreCase)
                && (conclusion.GetString() ?? string.Empty).Contains("setupplugin.plg", StringComparison.OrdinalIgnoreCase)
                && (conclusion.GetString() ?? string.Empty).Contains("plugin-provider.json", StringComparison.OrdinalIgnoreCase)
                && (conclusion.GetString() ?? string.Empty).Contains("entryDll", StringComparison.OrdinalIgnoreCase)
                && (conclusion.GetString() ?? string.Empty).Contains("%workingroot%/index.html", StringComparison.OrdinalIgnoreCase)
                && (conclusion.GetString() ?? string.Empty).Contains("kplugin", StringComparison.OrdinalIgnoreCase)
                && (conclusion.GetString() ?? string.Empty).Contains("DllRegisterServer", StringComparison.OrdinalIgnoreCase)
                && (conclusion.GetString() ?? string.Empty).Contains("GetExtensionJsApiObj", StringComparison.OrdinalIgnoreCase),
                "WPS reconnaissance should distinguish the shared COM add-ins UI path and API surface from the stronger field-level native package evidence that still leaves AddinsWl activation semantics unvalidated");
        }
    }

    private static void TestProbeOutputStructureIncludesTopLevelSupportSummary()
    {
        string core = ReadProjectSource("Installer.Core.ps1");

        AssertTrue(core.Contains("function Get-SupportSummary"),
            "probe core should define a top-level support summary builder");
        AssertTrue(core.Contains("DetectedHosts"),
            "probe output should summarize which hosts were detected");
        AssertTrue(core.Contains("SupportedHosts"),
            "probe output should summarize which detected hosts are currently confirmed supported");
        AssertTrue(core.Contains("PlannedHosts"),
            "probe output should summarize which detected hosts are still only planned");
        AssertTrue(core.Contains("UnmappedHosts"),
            "probe output should summarize which detected hosts are missing a support-matrix mapping");
    }

    private static void TestProbeOutputStructureIncludesMissingAndAmbiguousHostSummaries()
    {
        string core = ReadProjectSource("Installer.Core.ps1");

        AssertTrue(core.Contains("MissingExpectedHosts"),
            "probe output should summarize planned host combinations that are not detected on the current machine");
        AssertTrue(core.Contains("AmbiguousHosts"),
            "probe output should reserve a summary bucket for ambiguous host detections");
    }

    private static void TestProbeOutputStructureIncludesAmbiguityReasonDetails()
    {
        string core = ReadProjectSource("Installer.Core.ps1");

        AssertTrue(core.Contains("function Get-AmbiguityReason"),
            "probe core should define an ambiguity-reason resolver");
        AssertTrue(core.Contains("AmbiguityReason"),
            "probe output should expose a readable ambiguity reason for uncertain host detections");
        AssertTrue(core.Contains("Bitness could not be determined")
            || core.Contains("Registration view could not be determined"),
            "probe core should describe the reason when a host is classified as ambiguous");
    }

    private static void TestProbeSupportsSavingJsonOutputToFile()
    {
        string tempOutputPath = Path.Combine(Path.GetTempPath(), "wordtools-probe-" + Guid.NewGuid().ToString("N") + ".json");
        string scriptPath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Installer.Core.ps1"));

        try
        {
            var startInfo = new System.Diagnostics.ProcessStartInfo
            {
                FileName = "powershell",
                Arguments = "-NoProfile -ExecutionPolicy Bypass -File \"" + scriptPath + "\" -Mode Probe -OutputPath \"" + tempOutputPath + "\"",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            using var process = System.Diagnostics.Process.Start(startInfo);
            string standardOutput = process.StandardOutput.ReadToEnd();
            string standardError = process.StandardError.ReadToEnd();
            process.WaitForExit();

            AssertEqual(0, process.ExitCode, "probe should complete successfully when an output path is provided. stderr: " + standardError);
            AssertTrue(File.Exists(tempOutputPath), "probe should write a JSON file when -OutputPath is supplied");

            string savedJson = File.ReadAllText(tempOutputPath);
            AssertTrue(savedJson.Contains("\"SupportSummary\""), "saved probe file should contain the top-level support summary");
            AssertTrue(savedJson.Contains("\"Hosts\""), "saved probe file should contain host details");
            AssertTrue(standardOutput.Contains("\"ProbeMode\""), "probe should continue writing JSON to stdout even when saving to a file");
        }
        finally
        {
            if (File.Exists(tempOutputPath))
            {
                File.Delete(tempOutputPath);
            }
        }
    }

    private static void TestProbeSupportsWritingSummaryTextWithHostStateDetails()
    {
        string tempSummaryPath = Path.Combine(Path.GetTempPath(), "wordtools-probe-" + Guid.NewGuid().ToString("N") + ".txt");
        string scriptPath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Installer.Core.ps1"));

        try
        {
            var startInfo = new System.Diagnostics.ProcessStartInfo
            {
                FileName = "powershell",
                Arguments = "-NoProfile -ExecutionPolicy Bypass -File \"" + scriptPath + "\" -Mode Probe -SummaryTextPath \"" + tempSummaryPath + "\" -EvidenceLabel \"CurrentTurn-ProbeSummary-HostDetails\"",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            using var process = System.Diagnostics.Process.Start(startInfo);
            string standardOutput = process.StandardOutput.ReadToEnd();
            string standardError = process.StandardError.ReadToEnd();
            process.WaitForExit();

            AssertEqual(0, process.ExitCode, "probe should complete successfully when a summary path is provided. stderr: " + standardError + " stdout: " + standardOutput);
            AssertTrue(File.Exists(tempSummaryPath), "probe should write a summary text file when -SummaryTextPath is supplied");

            string summaryText = File.ReadAllText(tempSummaryPath);
            AssertTrue(summaryText.Contains("DetectedHosts:", StringComparison.Ordinal),
                "probe summary text should expose the detected hosts line");
            AssertTrue(summaryText.Contains("DetectedHostDetails:", StringComparison.Ordinal),
                "probe summary text should expose a detected-host details section for diagnostics");
            AssertTrue(summaryText.Contains("Word x64", StringComparison.Ordinal)
                && summaryText.Contains("InstallState=detected", StringComparison.Ordinal)
                && summaryText.Contains("VersionLine=", StringComparison.Ordinal)
                && summaryText.Contains("ValidationStage=formal p0 passed", StringComparison.Ordinal)
                && summaryText.Contains("DiagnosticsBundleId=", StringComparison.Ordinal),
                "probe summary text should include the current Word host label, install state, version line, validation stage, and diagnostics bundle identifier for traceable diagnostics");
        }
        finally
        {
            if (File.Exists(tempSummaryPath))
            {
                File.Delete(tempSummaryPath);
            }
        }
    }

    private static void TestProbeSupportsAttachingEvidenceLabel()
    {
        string tempOutputPath = Path.Combine(Path.GetTempPath(), "wordtools-probe-" + Guid.NewGuid().ToString("N") + ".json");
        string scriptPath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Installer.Core.ps1"));
        const string evidenceLabel = "Win64-Word64-LabA";

        try
        {
            var startInfo = new System.Diagnostics.ProcessStartInfo
            {
                FileName = "powershell",
                Arguments = "-NoProfile -ExecutionPolicy Bypass -File \"" + scriptPath + "\" -Mode Probe -OutputPath \"" + tempOutputPath + "\" -EvidenceLabel \"" + evidenceLabel + "\"",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            using var process = System.Diagnostics.Process.Start(startInfo);
            string standardOutput = process.StandardOutput.ReadToEnd();
            string standardError = process.StandardError.ReadToEnd();
            process.WaitForExit();

            AssertEqual(0, process.ExitCode, "probe should accept -EvidenceLabel without failing. stderr: " + standardError);
            AssertTrue(File.Exists(tempOutputPath), "probe should still save output when an evidence label is supplied");

            string savedJson = File.ReadAllText(tempOutputPath);
            AssertTrue(savedJson.Contains("\"EvidenceLabel\":  \"" + evidenceLabel + "\"")
                || savedJson.Contains("\"EvidenceLabel\":\"" + evidenceLabel + "\""),
                "saved probe file should preserve the supplied evidence label");
            AssertTrue(standardOutput.Contains(evidenceLabel),
                "stdout JSON should also include the supplied evidence label");
        }
        finally
        {
            if (File.Exists(tempOutputPath))
            {
                File.Delete(tempOutputPath);
            }
        }
    }

    private static void TestProbeSupportsAppendingEvidenceMarkdownRows()
    {
        string tempOutputPath = Path.Combine(Path.GetTempPath(), "wordtools-probe-" + Guid.NewGuid().ToString("N") + ".json");
        string tempMarkdownPath = Path.Combine(Path.GetTempPath(), "wordtools-probe-" + Guid.NewGuid().ToString("N") + ".md");
        string scriptPath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Installer.Core.ps1"));
        const string evidenceLabel = "Win64-Word64-LabA";

        try
        {
            File.WriteAllText(
                tempMarkdownPath,
                "# Host Detection Matrix" + Environment.NewLine + Environment.NewLine +
                "## Probe Evidence Log" + Environment.NewLine + Environment.NewLine +
                "| Evidence Label | Probed At (UTC) | Detected Hosts | Planned Hosts | Missing Expected Hosts | Ambiguous Hosts |" + Environment.NewLine +
                "| --- | --- | --- | --- | --- | --- |" + Environment.NewLine);

            var startInfo = new System.Diagnostics.ProcessStartInfo
            {
                FileName = "powershell",
                Arguments = "-NoProfile -ExecutionPolicy Bypass -File \"" + scriptPath + "\" -Mode Probe -OutputPath \"" + tempOutputPath + "\" -EvidenceLabel \"" + evidenceLabel + "\" -AppendEvidenceMarkdown -EvidenceMarkdownPath \"" + tempMarkdownPath + "\"",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            using var process = System.Diagnostics.Process.Start(startInfo);
            string standardOutput = process.StandardOutput.ReadToEnd();
            string standardError = process.StandardError.ReadToEnd();
            process.WaitForExit();

            AssertEqual(0, process.ExitCode, "probe should append markdown evidence without failing. stderr: " + standardError);

            string markdown = File.ReadAllText(tempMarkdownPath);
            AssertTrue(markdown.Contains("| " + evidenceLabel + " |"),
                "appended markdown should include the evidence label");
            AssertTrue(markdown.Contains("Word x64"),
                "appended markdown should summarize detected hosts");
            AssertTrue(markdown.Contains("| Word x64 | Word x64 |"),
                "appended markdown should record supported and planned host columns in the markdown row");
            AssertTrue(markdown.Contains("Word x86"),
                "appended markdown should summarize missing expected hosts");
            AssertTrue(markdown.Contains("## Pending Validation Matrix"),
                "markdown append should upgrade legacy files to include the pending validation matrix section");
            AssertTrue(markdown.IndexOf("## Pending Validation Matrix", StringComparison.Ordinal) < markdown.IndexOf("## Probe Evidence Log", StringComparison.Ordinal),
                "pending validation matrix should appear before the probe evidence log after upgrade");
            AssertTrue(standardOutput.Contains(evidenceLabel),
                "stdout JSON should still include the evidence label when markdown append is enabled");
        }
        finally
        {
            if (File.Exists(tempOutputPath))
            {
                File.Delete(tempOutputPath);
            }

            if (File.Exists(tempMarkdownPath))
            {
                File.Delete(tempMarkdownPath);
            }
        }
    }

    private static void TestEvidenceMarkdownScaffoldUsesValidationStageStructure()
    {
        string core = ReadProjectSource("Installer.Core.ps1");

        AssertTrue(core.Contains("## Validation Stages", StringComparison.Ordinal),
            "evidence markdown scaffolding should define the validation-stage vocabulary");
        AssertTrue(core.Contains("| Host | Bitness | Support state | Validation stage | Evidence |", StringComparison.Ordinal),
            "pending validation matrix scaffolding should use support-state and validation-stage columns");
        AssertTrue(core.Contains("| Evidence Label | Probed At (UTC) | Validation Stage | Detected Hosts | Supported Hosts | Planned Hosts | Missing Expected Hosts | Ambiguous Hosts |", StringComparison.Ordinal),
            "probe evidence log scaffolding should record validation stage alongside supported hosts");
    }

    private static void TestEvidenceMarkdownScaffoldSeedsCurrentSupportMatrixBaseline()
    {
        string core = ReadProjectSource("Installer.Core.ps1");

        AssertTrue(core.Contains("| Word | x64 | Supported | formal p0 passed |", StringComparison.Ordinal),
            "pending validation matrix scaffolding should seed Word x64 as the current formally supported baseline");
        AssertTrue(core.Contains("| Word | x86 | Planned | probe only |", StringComparison.Ordinal),
            "pending validation matrix scaffolding should keep Word x86 in the planned probe-only stage");
        AssertTrue(core.Contains("| WPS | x86 | Planned | ui failed |", StringComparison.Ordinal),
            "pending validation matrix scaffolding should seed WPS x86 with its current failed UI stage");
        AssertTrue(core.Contains("| WPS | x64 | Planned | probe only |", StringComparison.Ordinal),
            "pending validation matrix scaffolding should keep WPS x64 in the planned probe-only stage");
    }

    private static void TestHostDetectionMatrixUsesPendingAndEvidenceSections()
    {
        string markdown = ReadProjectSource(Path.Combine("docs", "installer", "host-detection-matrix.md"));

        AssertTrue(markdown.Contains("## Validation Stages"),
            "host detection matrix should document the validation-stage vocabulary used by phased acceptance");
        AssertTrue(markdown.Contains("## Pending Validation Matrix"),
            "host detection matrix should keep a dedicated pending-validation section");
        AssertTrue(markdown.Contains("## Probe Evidence Log"),
            "host detection matrix should include a dedicated auto-appended evidence log section");
        AssertTrue(markdown.Contains("| Evidence Label | Probed At (UTC) | Validation Stage | Detected Hosts | Supported Hosts | Planned Hosts | Missing Expected Hosts | Ambiguous Hosts |"),
            "probe evidence log should expose validation-stage and supported-hosts columns");
        AssertTrue(markdown.Contains("| Host | Bitness | Support state | Validation stage | Evidence |"),
            "pending validation matrix should expose support state and validation stage columns");
        AssertTrue(markdown.Contains("| Word | x64 | Supported | formal p0 passed |"),
            "pending validation matrix should mark the locally verified Word x64 target as supported with formal P0 evidence");
        AssertTrue(markdown.Contains("| WPS | x86 | Planned | ui failed |"),
            "pending validation matrix should keep WPS x86 planned while recording the failed UI stage");
        AssertTrue(markdown.Contains("CurrentMachine-WpsX86-UiLoadFailure-20260524"),
            "host detection matrix should record the failed WPS UI evidence that triggered the rollback");
    }

    private static void TestPlanModeExposesDryRunRegistrationPlanForSupportedWordX64()
    {
        string scriptPath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Installer.Core.ps1"));

        var startInfo = new System.Diagnostics.ProcessStartInfo
        {
            FileName = "powershell",
            Arguments = "-NoProfile -ExecutionPolicy Bypass -File \"" + scriptPath + "\" -Mode Plan -EvidenceLabel \"CurrentMachine-Word64-20260517\"",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        using var process = System.Diagnostics.Process.Start(startInfo);
        string standardOutput = process.StandardOutput.ReadToEnd();
        string standardError = process.StandardError.ReadToEnd();
        process.WaitForExit();

        AssertEqual(0, process.ExitCode, "plan mode should complete successfully. stderr: " + standardError);
        AssertTrue(standardOutput.Contains("\"RegistrationPlan\""),
            "plan mode should emit a top-level registration plan payload");

        using JsonDocument json = JsonDocument.Parse(standardOutput);
        JsonElement registrationPlan = json.RootElement.GetProperty("RegistrationPlan");
        JsonElement planSummary = registrationPlan.GetProperty("PlanSummary");

        AssertTrue(standardOutput.Contains("\"ExecutionMode\":  \"DryRun\"")
            || standardOutput.Contains("\"ExecutionMode\":\"DryRun\""),
            "plan mode should make it explicit that no registration is actually performed");
        AssertTrue(standardOutput.Contains("WordTools.ThisAddIn"),
            "plan mode should reveal the future add-in ProgId for Word");
        AssertTrue(standardOutput.Contains("Microsoft\\\\Office\\\\Word\\\\Addins")
            || standardOutput.Contains("Microsoft\\Office\\Word\\Addins"),
            "plan mode should reveal the future Word add-in registry path");
        AssertTrue(standardOutput.Contains("RegAsm.exe"),
            "plan mode should reveal which RegAsm path would be used");
        AssertEqual("DryRun", planSummary.GetProperty("ExecutionMode").GetString(),
            "plan summary should remain explicit about dry-run execution");
        AssertEqual(1, planSummary.GetProperty("DetectedTargetCount").GetInt32(),
            "current machine should expose one detected target in the plan summary");
        AssertEqual(1, planSummary.GetProperty("RegistrableTargetCount").GetInt32(),
            "current machine should expose one registrable target in the plan summary");
        AssertEqual(0, planSummary.GetProperty("SkippedTargetCount").GetInt32(),
            "current machine should not mark the supported Word x64 host as skipped");

        string[] requiredRegAsmModes = planSummary.GetProperty("RequiredRegAsmModes")
            .EnumerateArray()
            .Select(static item => item.GetString())
            .Where(static item => !string.IsNullOrWhiteSpace(item))
            .ToArray();
        AssertTrue(requiredRegAsmModes.SequenceEqual(new[] { "x64" }, StringComparer.Ordinal),
            "current machine should require only x64 RegAsm in dry-run planning");

        string[] registryWrites = planSummary.GetProperty("RegistryWrites")
            .EnumerateArray()
            .Select(static item => item.GetString())
            .Where(static item => !string.IsNullOrWhiteSpace(item))
            .ToArray();
        AssertTrue(registryWrites.Any(static item => item.Contains(@"Microsoft\Office\Word\Addins\WordTools.ThisAddIn", StringComparison.Ordinal)),
            "plan summary should include the future Word add-in registry write");

        string[] plannedActionUnion = planSummary.GetProperty("PlannedActionUnion")
            .EnumerateArray()
            .Select(static item => item.GetString())
            .Where(static item => !string.IsNullOrWhiteSpace(item))
            .ToArray();
        AssertTrue(plannedActionUnion.Contains("RegisterComCodebase", StringComparer.Ordinal),
            "plan summary should include COM registration in the action union");
        AssertTrue(plannedActionUnion.Contains("WriteWordAddInRegistry", StringComparer.Ordinal),
            "plan summary should include Word add-in registry writes in the action union");

        string overallDecision = planSummary.GetProperty("OverallDecision").GetString() ?? string.Empty;
        AssertTrue(overallDecision.Contains("dry-run", StringComparison.OrdinalIgnoreCase)
            && overallDecision.Contains("without executing any live registration", StringComparison.OrdinalIgnoreCase),
            "plan summary should explicitly state that the overall decision remains dry-run only");
    }

    private static void TestPlanModeRespectsExplicitWordX86RequestWhenNoWordX86HostIsDetected()
    {
        string scriptPath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Installer.Core.ps1"));

        var startInfo = new System.Diagnostics.ProcessStartInfo
        {
            FileName = "powershell",
            Arguments =
                "-NoProfile -ExecutionPolicy Bypass -File \"" + scriptPath + "\""
                + " -Mode Plan"
                + " -RequestedHost Word"
                + " -Architecture x86"
                + " -EvidenceLabel \"CurrentTurn-WordX86PlanRegression\"",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        using var process = System.Diagnostics.Process.Start(startInfo);
        string standardOutput = process.StandardOutput.ReadToEnd();
        string standardError = process.StandardError.ReadToEnd();
        process.WaitForExit();

        AssertEqual(0, process.ExitCode, "word x86 plan mode should complete successfully even when no Word x86 host is detected. stderr: " + standardError);

        using JsonDocument json = JsonDocument.Parse(standardOutput);
        JsonElement planSummary = json.RootElement.GetProperty("RegistrationPlan").GetProperty("PlanSummary");
        JsonElement targets = json.RootElement.GetProperty("RegistrationPlan").GetProperty("Targets");

        AssertEqual(0, planSummary.GetProperty("DetectedTargetCount").GetInt32(),
            "explicit Word x86 dry-run planning should not borrow the current Word x64 host into the filtered target set");
        AssertEqual(0, planSummary.GetProperty("RegistrableTargetCount").GetInt32(),
            "explicit Word x86 dry-run planning should expose zero registrable targets when no Word x86 host is detected");
        AssertEqual(0, planSummary.GetProperty("SkippedTargetCount").GetInt32(),
            "explicit Word x86 dry-run planning should not treat the current Word x64 host as a skipped x86 target");
        AssertEqual(0, targets.GetArrayLength(),
            "explicit Word x86 dry-run planning should emit no per-target payload when the requested host and architecture are absent");
        AssertTrue((planSummary.GetProperty("OverallDecision").GetString() ?? string.Empty).Contains("Would not register any detected host", StringComparison.Ordinal),
            "explicit Word x86 dry-run planning should explain that no live-registration-eligible target matched the current request");
    }

    private static void TestPlanModeSummaryTextExposesHostStateDetailsForDiagnostics()
    {
        string scriptPath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Installer.Core.ps1"));
        string tempSummaryPath = Path.Combine(Path.GetTempPath(), "wordtools-plan-summary-" + Guid.NewGuid().ToString("N") + ".txt");

        try
        {
            var startInfo = new System.Diagnostics.ProcessStartInfo
            {
                FileName = "powershell",
                Arguments =
                    "-NoProfile -ExecutionPolicy Bypass -File \"" + scriptPath + "\""
                    + " -Mode Plan"
                    + " -EvidenceLabel \"CurrentTurn-PlanSummary-HostDetails\""
                    + " -SummaryTextPath \"" + tempSummaryPath + "\"",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            using var process = System.Diagnostics.Process.Start(startInfo);
            string standardOutput = process.StandardOutput.ReadToEnd();
            string standardError = process.StandardError.ReadToEnd();
            process.WaitForExit();

            AssertEqual(0, process.ExitCode, "plan mode should complete successfully when writing a summary text file. stderr: " + standardError + " stdout: " + standardOutput);
            AssertTrue(File.Exists(tempSummaryPath),
                "plan mode should persist a summary text file when -SummaryTextPath is supplied");

            string summaryText = File.ReadAllText(tempSummaryPath);
            AssertTrue(summaryText.Contains("DetectedHostDetails:", StringComparison.Ordinal),
                "plan summary text should expose a detected-host details section for diagnostics");
            AssertTrue(summaryText.Contains("Word x64", StringComparison.Ordinal)
                && summaryText.Contains("InstallState=detected", StringComparison.Ordinal)
                && summaryText.Contains("VersionLine=", StringComparison.Ordinal)
                && summaryText.Contains("ValidationStage=formal p0 passed", StringComparison.Ordinal)
                && summaryText.Contains("DiagnosticsBundleId=", StringComparison.Ordinal),
                "plan summary text should include the current Word host label, install state, version line, validation stage, and diagnostics bundle identifier for traceable diagnostics");
        }
        finally
        {
            if (File.Exists(tempSummaryPath))
            {
                File.Delete(tempSummaryPath);
            }
        }
    }

    private static void TestUnregisterPreviewRespectsExplicitWordX86RequestWhenNoWordX86HostIsDetected()
    {
        string scriptPath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Installer.Core.ps1"));

        var startInfo = new System.Diagnostics.ProcessStartInfo
        {
            FileName = "powershell",
            Arguments =
                "-NoProfile -ExecutionPolicy Bypass -File \"" + scriptPath + "\""
                + " -Mode Unregister"
                + " -RequestedHost Word"
                + " -Architecture x86"
                + " -EvidenceLabel \"CurrentTurn-WordX86UnregisterRegression\"",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        using var process = System.Diagnostics.Process.Start(startInfo);
        string standardOutput = process.StandardOutput.ReadToEnd();
        string standardError = process.StandardError.ReadToEnd();
        process.WaitForExit();

        AssertEqual(0, process.ExitCode, "word x86 unregister preview should complete successfully even when no Word x86 host is detected. stderr: " + standardError);

        using JsonDocument json = JsonDocument.Parse(standardOutput);
        JsonElement unregisterPlan = json.RootElement.GetProperty("UnregisterPlan");
        JsonElement previewSummary = unregisterPlan.GetProperty("UnregisterPreviewSummary");
        JsonElement installerHandoffSummary = unregisterPlan.GetProperty("InstallerHandoffSummary");
        JsonElement installerPreviewReport = unregisterPlan.GetProperty("InstallerPreviewReport");
        JsonElement operationManifest = unregisterPlan.GetProperty("OperationManifest");
        JsonElement targets = unregisterPlan.GetProperty("Targets");

        AssertEqual(0, unregisterPlan.GetProperty("PreviewableCount").GetInt32(),
            "explicit Word x86 unregister preview should expose zero previewable targets when no Word x86 host is detected");
        AssertEqual(0, previewSummary.GetProperty("DetectedTargetCount").GetInt32(),
            "explicit Word x86 unregister preview should not borrow the current Word x64 host into the filtered target set");
        AssertEqual(0, previewSummary.GetProperty("PreviewableTargetCount").GetInt32(),
            "explicit Word x86 unregister preview should expose zero eligible targets when the requested host and architecture are absent");
        AssertEqual(0, operationManifest.GetArrayLength(),
            "explicit Word x86 unregister preview should emit no operation manifest entries when no matching host is detected");
        AssertEqual(0, targets.GetArrayLength(),
            "explicit Word x86 unregister preview should emit no per-target payload when the requested host and architecture are absent");
        AssertEqual(0, installerHandoffSummary.GetProperty("DetectedHostCount").GetInt32(),
            "explicit Word x86 unregister preview should scope installer handoff diagnostics to the filtered request instead of the ambient Word x64 host");
        AssertEqual(0, installerHandoffSummary.GetProperty("SupportedHostCount").GetInt32(),
            "explicit Word x86 unregister preview should not report the ambient Word x64 host as a supported handoff target for an x86-only request");
        string previewReportSummary = installerPreviewReport.GetProperty("SummaryText").GetString() ?? string.Empty;
        AssertTrue(previewReportSummary.Contains("Detected hosts: 0.", StringComparison.Ordinal)
            && previewReportSummary.Contains("Supported hosts: none.", StringComparison.Ordinal)
            && previewReportSummary.Contains("Preview actions: 0.", StringComparison.Ordinal),
            "explicit Word x86 unregister preview report should summarize the filtered request rather than the ambient Word x64 host");
        AssertTrue((previewSummary.GetProperty("OverallDecision").GetString() ?? string.Empty).Contains("found no live-unregistration-eligible targets", StringComparison.OrdinalIgnoreCase),
            "explicit Word x86 unregister preview should explain that no live-unregistration-eligible target matched the current request");
    }

    private static void TestDryRunSummaryTextDerivesRegistrableTargetsFromDryRunEligibility()
    {
        string core = ReadProjectSource("Installer.Core.ps1");

        AssertTrue(core.Contains("$DryRunPlan.Targets |", StringComparison.Ordinal)
            && core.Contains("Where-Object { $_.WouldRegister }", StringComparison.Ordinal)
            && core.Contains("ForEach-Object { \"{0} {1}\" -f $_.HostName, $_.HostBitness }", StringComparison.Ordinal),
            "dry-run decision summary text should derive registrable targets from dry-run eligibility rather than mirroring the supported-target label list");
        AssertTrue(core.IndexOf("$registrableTargets = @($DryRunPlan.PlanSummary.SupportedTargetLabels)", StringComparison.Ordinal) < 0,
            "dry-run decision summary text should stop wiring registrable targets directly to supported target labels");
    }

    private static void TestUnregisterPreviewSummaryTextExposesHostStateDetailsForDiagnostics()
    {
        string scriptPath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Installer.Core.ps1"));
        string tempSummaryPath = Path.Combine(Path.GetTempPath(), "wordtools-unregister-summary-" + Guid.NewGuid().ToString("N") + ".txt");

        try
        {
            var startInfo = new System.Diagnostics.ProcessStartInfo
            {
                FileName = "powershell",
                Arguments =
                    "-NoProfile -ExecutionPolicy Bypass -File \"" + scriptPath + "\""
                    + " -Mode Unregister"
                    + " -EvidenceLabel \"CurrentTurn-UnregisterSummary-HostDetails\""
                    + " -SummaryTextPath \"" + tempSummaryPath + "\"",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            using var process = System.Diagnostics.Process.Start(startInfo);
            string standardOutput = process.StandardOutput.ReadToEnd();
            string standardError = process.StandardError.ReadToEnd();
            process.WaitForExit();

            AssertEqual(0, process.ExitCode, "unregister preview mode should complete successfully when writing a summary text file. stderr: " + standardError + " stdout: " + standardOutput);
            AssertTrue(File.Exists(tempSummaryPath),
                "unregister preview mode should persist a summary text file when -SummaryTextPath is supplied");

            string summaryText = File.ReadAllText(tempSummaryPath);
            AssertTrue(summaryText.Contains("DetectedHostDetails:", StringComparison.Ordinal),
                "unregister preview summary text should expose a detected-host details section for diagnostics");
            AssertTrue(summaryText.Contains("Word x64", StringComparison.Ordinal)
                && summaryText.Contains("InstallState=detected", StringComparison.Ordinal)
                && summaryText.Contains("VersionLine=", StringComparison.Ordinal)
                && summaryText.Contains("ValidationStage=formal p0 passed", StringComparison.Ordinal)
                && summaryText.Contains("DiagnosticsBundleId=", StringComparison.Ordinal),
                "unregister preview summary text should include the current Word host label, install state, version line, validation stage, and diagnostics bundle identifier for traceable diagnostics");
        }
        finally
        {
            if (File.Exists(tempSummaryPath))
            {
                File.Delete(tempSummaryPath);
            }
        }
    }

    private static void TestLiveRegisterModeFailsFastWithExplicitAdministratorGuidance()
    {
        string scriptPath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Installer.Core.ps1"));

        var startInfo = new System.Diagnostics.ProcessStartInfo
        {
            FileName = "powershell",
            Arguments =
                "-NoProfile -ExecutionPolicy Bypass -File \"" + scriptPath + "\""
                + " -Mode Register"
                + " -ExecutionIntent Live"
                + " -RequestedHost Word"
                + " -Architecture Auto"
                + " -Configuration Release"
                + " -DllPathOverride \"" + scriptPath + "\""
                + " -EvidenceLabel \"AdminPreflightTest\"",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        using var process = System.Diagnostics.Process.Start(startInfo);
        string standardOutput = process.StandardOutput.ReadToEnd();
        string standardError = process.StandardError.ReadToEnd();
        process.WaitForExit();

        AssertTrue(process.ExitCode != 0,
            "live register mode should fail in the current non-elevated test environment");

        string combinedOutput = standardOutput + Environment.NewLine + standardError;
        AssertTrue(combinedOutput.Contains("elevated administrator PowerShell session", StringComparison.OrdinalIgnoreCase),
            "live register mode should fail with explicit administrator guidance before attempting real COM registration");
        AssertTrue(!combinedOutput.Contains("RegAsm64 failed", StringComparison.OrdinalIgnoreCase)
            && !combinedOutput.Contains("RA0000", StringComparison.OrdinalIgnoreCase),
            "live register mode should fail before invoking RegAsm when elevation is missing");
    }

    private static void TestLiveRegisterModeWritesFailurePayloadAndSummaryWhenLiveExecutionFails()
    {
        string scriptPath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Installer.Core.ps1"));
        string tempOutputPath = Path.Combine(Path.GetTempPath(), "wordtools-live-register-failure-" + Guid.NewGuid().ToString("N") + ".json");
        string tempSummaryPath = Path.Combine(Path.GetTempPath(), "wordtools-live-register-failure-" + Guid.NewGuid().ToString("N") + ".txt");

        try
        {
            var startInfo = new System.Diagnostics.ProcessStartInfo
            {
                FileName = "powershell",
                Arguments =
                    "-NoProfile -ExecutionPolicy Bypass -File \"" + scriptPath + "\""
                    + " -Mode Register"
                    + " -ExecutionIntent Live"
                    + " -RequestedHost Word"
                    + " -Architecture Auto"
                    + " -Configuration Release"
                    + " -DllPathOverride \"" + scriptPath + "\""
                    + " -EvidenceLabel \"AdminFailurePayloadTest\""
                    + " -OutputPath \"" + tempOutputPath + "\""
                    + " -SummaryTextPath \"" + tempSummaryPath + "\"",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            using var process = System.Diagnostics.Process.Start(startInfo);
            string standardOutput = process.StandardOutput.ReadToEnd();
            string standardError = process.StandardError.ReadToEnd();
            process.WaitForExit();

            AssertTrue(process.ExitCode != 0,
                "live register mode should still fail in the current non-elevated test environment");
            AssertTrue(File.Exists(tempOutputPath),
                "live register mode should persist a failure JSON payload when -OutputPath is supplied");
            AssertTrue(File.Exists(tempSummaryPath),
                "live register mode should persist a failure summary when -SummaryTextPath is supplied");

            string savedJson = File.ReadAllText(tempOutputPath);
            using JsonDocument document = JsonDocument.Parse(savedJson);
            JsonElement liveFailure = document.RootElement.GetProperty("LiveFailure");

            AssertEqual("Register", liveFailure.GetProperty("Operation").GetString(),
                "failure payload should preserve the attempted live operation");
            AssertEqual("Live", liveFailure.GetProperty("ExecutionMode").GetString(),
                "failure payload should mark the failed path as live execution");
            AssertEqual(false, liveFailure.GetProperty("Succeeded").GetBoolean(),
                "failure payload should explicitly mark the execution as failed");
            AssertTrue((liveFailure.GetProperty("ErrorMessage").GetString() ?? string.Empty)
                .Contains("elevated administrator PowerShell session", StringComparison.OrdinalIgnoreCase),
                "failure payload should preserve the administrator guidance message");

            string summaryText = File.ReadAllText(tempSummaryPath);
            AssertTrue(summaryText.Contains("Shared installer core failed during live Register.", StringComparison.Ordinal),
                "failure summary should identify the failed live operation");
            AssertTrue(summaryText.Contains("elevated administrator PowerShell session", StringComparison.OrdinalIgnoreCase),
                "failure summary should preserve the administrator guidance message");

            string combinedOutput = standardOutput + Environment.NewLine + standardError;
            AssertTrue(combinedOutput.Contains("elevated administrator PowerShell session", StringComparison.OrdinalIgnoreCase),
                "live register mode should continue surfacing the administrator guidance in console output");
        }
        finally
        {
            if (File.Exists(tempOutputPath))
            {
                File.Delete(tempOutputPath);
            }

            if (File.Exists(tempSummaryPath))
            {
                File.Delete(tempSummaryPath);
            }
        }
    }

    private static void TestSharedCoreSourceSupportsInstallerDrivenSelfElevation()
    {
        string source = ReadProjectSource("Installer.Core.ps1");

        AssertTrue(source.Contains("[switch]$AllowSelfElevation", StringComparison.Ordinal)
            && source.Contains("[switch]$LiveElevatedRelaunch", StringComparison.Ordinal),
            "shared installer core should expose the self-elevation guard switches");
        AssertTrue(source.Contains("function Invoke-SelfElevatedLiveExecution", StringComparison.Ordinal)
            && source.Contains("Start-Process -FilePath \"powershell.exe\" -Verb RunAs", StringComparison.Ordinal),
            "shared installer core should define an installer-driven self-elevation helper");
        AssertTrue(source.Contains("Invoke-SelfElevatedLiveExecution -Operation \"Register\"", StringComparison.Ordinal)
            && source.Contains("Invoke-SelfElevatedLiveExecution -Operation \"Unregister\"", StringComparison.Ordinal),
            "live register and unregister flows should reuse the self-elevation helper before real COM registration");
    }

    private static void TestSharedCoreSourceCapturesExternalToolOutputBeforeReturningLiveObjects()
    {
        string source = ReadProjectSource("Installer.Core.ps1");

        AssertTrue(source.Contains("function Invoke-NativeToolCapture", StringComparison.Ordinal)
            && source.Contains("-RedirectStandardOutput", StringComparison.Ordinal)
            && source.Contains("-RedirectStandardError", StringComparison.Ordinal),
            "shared installer core should capture native tool stdout and stderr explicitly instead of leaking them into the function return pipeline");
        AssertTrue(source.Contains("ToolOutput    = $toolInvocation.CombinedOutput", StringComparison.Ordinal),
            "shared installer core should preserve captured RegAsm/NGen output inside the returned live execution objects");
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
