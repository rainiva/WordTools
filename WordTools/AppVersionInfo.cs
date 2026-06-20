using System.Reflection;

namespace WordTools
{
    internal static class AppVersionInfo
    {
        /// <summary>
        /// 显示用版本号，优先读取 sync-version.ps1 写入的 AssemblyInformationalVersion（x.x.x）。
        /// </summary>
        public static string DisplayVersion
        {
            get
            {
                Assembly assembly = Assembly.GetExecutingAssembly();
                string informational = assembly
                    .GetCustomAttribute<AssemblyInformationalVersionAttribute>()
                    ?.InformationalVersion;

                if (!string.IsNullOrWhiteSpace(informational))
                {
                    // 去掉 +build / -pre 等 semver 后缀，仅保留 x.x.x
                    int cut = informational.IndexOfAny(new[] { '+', '-' });
                    return cut >= 0 ? informational.Substring(0, cut) : informational;
                }

                System.Version version = assembly.GetName().Version;
                return version == null ? "unknown" : version.ToString(3);
            }
        }

        public static string AboutMessage
        {
            get
            {
                Assembly assembly = Assembly.GetExecutingAssembly();
                string copyright = assembly.GetCustomAttribute<AssemblyCopyrightAttribute>()?.Copyright;
                string copyrightLine = string.IsNullOrWhiteSpace(copyright)
                    ? string.Empty
                    : copyright.Replace("Copyright ", string.Empty).Trim();

                return string.Format(
                    "Word工具箱 v{0}\n\n功能：批量插图、Excel数据填充、自动编号\n\n{1}",
                    DisplayVersion,
                    copyrightLine);
            }
        }
    }
}
