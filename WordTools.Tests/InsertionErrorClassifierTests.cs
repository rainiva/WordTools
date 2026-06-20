using System;
using System.IO;
using WordTools.Services;
using Xunit;

namespace WordTools.Tests
{
    public class InsertionErrorClassifierTests
    {
        [Fact]
        public void Classify_NullException_ReturnsUnknown()
        {
            Assert.Equal("未知错误", InsertionErrorClassifier.Classify(null));
        }

        [Fact]
        public void Classify_FileNotFound_ReturnsFileMissing()
        {
            var ex = new FileNotFoundException("找不到文件");
            Assert.Equal("文件不存在或已被移动", InsertionErrorClassifier.Classify(ex));
        }

        [Fact]
        public void Classify_IOException_ReturnsFileLocked()
        {
            var ex = new IOException("进程无法访问该文件");
            Assert.Equal("文件被其他程序占用", InsertionErrorClassifier.Classify(ex));
        }

        [Fact]
        public void Classify_UnauthorizedAccess_ReturnsAccessDenied()
        {
            var ex = new UnauthorizedAccessException("拒绝访问");
            Assert.Equal("没有文件访问权限", InsertionErrorClassifier.Classify(ex));
        }

        [Fact]
        public void IsMergedCellError_MergeKeyword_ReturnsTrue()
        {
            var ex = new Exception("合并单元格操作失败");
            Assert.True(InsertionErrorClassifier.IsMergedCellError(ex));
        }

        [Fact]
        public void IsMergedCellError_NormalError_ReturnsFalse()
        {
            var ex = new Exception("普通错误");
            Assert.False(InsertionErrorClassifier.IsMergedCellError(ex));
        }
    }
}
