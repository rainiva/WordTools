using System.Reflection;

namespace WordTools
{
    internal static class AppVersionInfo
    {
        public static string DisplayVersion
        {
            get
            {
                System.Version version = Assembly.GetExecutingAssembly().GetName().Version;
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
