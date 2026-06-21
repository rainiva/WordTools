using System;
using System.IO;
using System.Linq;
using System.Reflection;
using Xunit;

namespace WordTools.Tests
{
    public class ThisAddInAutomationServiceTests
    {
        private const string AutomationServiceInterfaceName = "IRequestComAddInAutomationService";

        [Fact]
        public void ThisAddIn_exposes_com_automation_service_for_ui_e2e()
        {
            var dllPath = ResolveWordToolsDll();
            var assembly = Assembly.LoadFrom(dllPath);
            var addInType = assembly.GetType("WordTools.ThisAddIn", throwOnError: true);

            Assert.Contains(
                AutomationServiceInterfaceName,
                addInType.GetInterfaces().Select(type => type.Name));

            var method = addInType.GetMethod("GetComAddInAutomationService");
            Assert.NotNull(method);
            Assert.NotNull(addInType.GetMethod("Automation_ShowInsertPhotosForm"));
            Assert.NotNull(addInType.GetMethod("Automation_ExecuteFromConfig"));

            var instance = Activator.CreateInstance(addInType);
            var automationObject = method.Invoke(instance, null);
            Assert.Same(instance, automationObject);
        }

        private static string ResolveWordToolsDll()
        {
            var current = new DirectoryInfo(AppContext.BaseDirectory);
            while (current != null)
            {
                var candidate = Path.Combine(current.FullName, "WordTools", "bin", "Release", "WordTools.dll");
                if (File.Exists(candidate))
                {
                    return candidate;
                }

                current = current.Parent;
            }

            throw new FileNotFoundException("WordTools Release DLL not found. Build WordTools before running this test.");
        }
    }
}
