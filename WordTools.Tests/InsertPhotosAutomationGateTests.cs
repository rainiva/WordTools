using System;
using WordTools.Forms;
using WordTools.Services;
using Xunit;
using System.IO;

namespace WordTools.Tests
{
    public class InsertPhotosAutomationGateTests
    {
        [Fact]
        public void IsEnabled_returns_true_when_env_is_one()
        {
            Environment.SetEnvironmentVariable(InsertPhotosAutomationGate.EnableEnvVar, "1");
            try
            {
                Assert.True(InsertPhotosAutomationGate.IsEnabled);
            }
            finally
            {
                Environment.SetEnvironmentVariable(InsertPhotosAutomationGate.EnableEnvVar, null);
            }
        }

        [Fact]
        public void IsEnabled_returns_false_when_env_missing()
        {
            Environment.SetEnvironmentVariable(InsertPhotosAutomationGate.EnableEnvVar, null);
            Assert.False(InsertPhotosAutomationGate.IsEnabled);
        }

        [Fact]
        public void GetPresetSelectedFiles_splits_semicolon_paths()
        {
            Environment.SetEnvironmentVariable(
                InsertPhotosAutomationGate.SelectedFilesEnvVar,
                @"C:\a\01.jpg;C:\b\02.jpg");
            try
            {
                var files = InsertPhotosAutomationGate.GetPresetSelectedFiles();
                Assert.Equal(2, files.Length);
                Assert.Equal(@"C:\a\01.jpg", files[0]);
                Assert.Equal(@"C:\b\02.jpg", files[1]);
            }
            finally
            {
                Environment.SetEnvironmentVariable(InsertPhotosAutomationGate.SelectedFilesEnvVar, null);
            }
        }

        [Fact]
        public void EnsureEnabled_throws_when_not_enabled()
        {
            Environment.SetEnvironmentVariable(InsertPhotosAutomationGate.EnableEnvVar, null);
            Assert.Throws<InvalidOperationException>(() => InsertPhotosAutomationGate.EnsureEnabled());
        }

        [Fact]
        public void ResolveFormPreset_b03_uses_filename_and_numbering()
        {
            var preset = InsertPhotosAutomationGate.ResolveFormPreset("AC-UI-B03");
            Assert.True(preset.UseFileNameAsDescription);
            Assert.True(preset.NeedAutoNumbering);
        }

        [Fact]
        public void ResolveFormPreset_b05_disables_numbering()
        {
            var preset = InsertPhotosAutomationGate.ResolveFormPreset("AC-UI-B05");
            Assert.True(preset.UseFileNameAsDescription);
            Assert.False(preset.NeedAutoNumbering);
        }

        [Fact]
        public void ResolveFormPreset_b07_root_only()
        {
            var preset = InsertPhotosAutomationGate.ResolveFormPreset("AC-UI-B07");
            Assert.True(preset.IncludeRootImages);
            Assert.False(preset.IncludeSubFolderImages);
        }

        [Fact]
        public void ResolveFormPreset_b09_no_description()
        {
            var preset = InsertPhotosAutomationGate.ResolveFormPreset("AC-UI-B09");
            Assert.False(preset.UseFileNameAsDescription);
            Assert.False(preset.NeedDescription);
            Assert.False(preset.NeedAutoNumbering);
        }

        [Fact]
        public void ResolveFormPreset_b10_number_after_center()
        {
            var preset = InsertPhotosAutomationGate.ResolveFormPreset("AC-UI-B10");
            Assert.Equal(2, preset.NumberPosition);
            Assert.Equal(2, preset.NumberAlignment);
        }

        [Fact]
        public void ResolveFormPreset_b11_folder_name()
        {
            var preset = InsertPhotosAutomationGate.ResolveFormPreset("AC-UI-B11");
            Assert.True(preset.UseFolderNameAsDescription);
            Assert.False(preset.UseFileNameAsDescription);
        }

        [Fact]
        public void ResolveFormPreset_b15_fixed_height()
        {
            var preset = InsertPhotosAutomationGate.ResolveFormPreset("AC-UI-B15");
            Assert.True(preset.MinHeightCm.HasValue);
            Assert.Equal(3f, preset.MinHeightCm.Value);
        }

        [Fact]
        public void GetCaseId_reads_from_config_file_when_env_points_to_file()
        {
            var configPath = Path.Combine(Path.GetTempPath(), "wordtools-gate-test-" + Guid.NewGuid().ToString("N") + ".json");
            File.WriteAllText(configPath, "{\"case_id\":\"AC-UI-B07\",\"folder_path\":\"C:\\\\temp\\\\test2\"}");
            Environment.SetEnvironmentVariable(InsertPhotosAutomationGate.ConfigFileEnvVar, configPath);
            try
            {
                Assert.Equal("AC-UI-B07", InsertPhotosAutomationGate.GetCaseId());
                Assert.Equal("C:\\temp\\test2", InsertPhotosAutomationGate.GetPresetFolderPath());
            }
            finally
            {
                Environment.SetEnvironmentVariable(InsertPhotosAutomationGate.ConfigFileEnvVar, null);
                File.Delete(configPath);
            }
        }

        [Fact]
        public void TryBuildRequest_builds_selected_files_for_b05_config()
        {
            Environment.SetEnvironmentVariable(InsertPhotosAutomationGate.EnableEnvVar, "1");
            var configPath = Path.Combine(Path.GetTempPath(), "wordtools-gate-req-" + Guid.NewGuid().ToString("N") + ".json");
            File.WriteAllText(
                configPath,
                "{\"case_id\":\"AC-UI-B05\",\"selected_files\":\"C:\\\\imgs\\\\a.jpg\"}");
            Environment.SetEnvironmentVariable(InsertPhotosAutomationGate.ConfigFileEnvVar, configPath);
            try
            {
                Assert.True(InsertPhotosAutomationGate.TryBuildRequest(out var request));
                Assert.Equal(InsertPhotosRequestMode.SelectedFiles, request.Mode);
                Assert.Single(request.SelectedFiles);
                Assert.Equal(@"C:\imgs\a.jpg", request.SelectedFiles[0]);
                Assert.False(request.NeedAutoNumbering);
                Assert.True(request.UseFileNameAsDescription);
            }
            finally
            {
                Environment.SetEnvironmentVariable(InsertPhotosAutomationGate.EnableEnvVar, null);
                Environment.SetEnvironmentVariable(InsertPhotosAutomationGate.ConfigFileEnvVar, null);
                File.Delete(configPath);
            }
        }

        [Fact]
        public void TryBuildRequest_builds_folder_root_only_for_b07_config()
        {
            Environment.SetEnvironmentVariable(InsertPhotosAutomationGate.EnableEnvVar, "1");
            var folder = Path.GetTempPath();
            var configPath = Path.Combine(Path.GetTempPath(), "wordtools-gate-req-" + Guid.NewGuid().ToString("N") + ".json");
            File.WriteAllText(
                configPath,
                "{\"case_id\":\"AC-UI-B07\",\"folder_path\":\"" + folder.Replace("\\", "\\\\") + "\"}");
            Environment.SetEnvironmentVariable(InsertPhotosAutomationGate.ConfigFileEnvVar, configPath);
            try
            {
                Assert.True(InsertPhotosAutomationGate.TryBuildRequest(out var request));
                Assert.Equal(InsertPhotosRequestMode.Folder, request.Mode);
                Assert.Equal(folder, request.FolderPath);
                Assert.True(request.IncludeRootImages);
                Assert.False(request.IncludeSubFolderImages);
            }
            finally
            {
                Environment.SetEnvironmentVariable(InsertPhotosAutomationGate.EnableEnvVar, null);
                Environment.SetEnvironmentVariable(InsertPhotosAutomationGate.ConfigFileEnvVar, null);
                File.Delete(configPath);
            }
        }

        [Fact]
        public void TryGetPresetFolderPath_returns_true_when_folder_exists()
        {
            var folder = Path.GetTempPath();
            Environment.SetEnvironmentVariable(InsertPhotosAutomationGate.FolderPathEnvVar, folder);
            try
            {
                Assert.True(InsertPhotosAutomationGate.TryGetPresetFolderPath(out var resolved));
                Assert.Equal(folder, resolved);
            }
            finally
            {
                Environment.SetEnvironmentVariable(InsertPhotosAutomationGate.FolderPathEnvVar, null);
            }
        }
    }
}
