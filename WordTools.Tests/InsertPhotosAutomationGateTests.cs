using System;
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
