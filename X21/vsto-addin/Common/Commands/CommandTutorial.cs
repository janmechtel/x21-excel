using X21.Common.Commands;
using X21.Common.Data;
using X21.Common.Model;
using X21.Utils;
using System;
using System.IO;
using System.Reflection;
using Microsoft.Office.Interop.Excel;

namespace X21.Common.Commands
{
    /// <summary>
    /// Command to open the tutorial samples file
    /// </summary>
    public class CommandTutorial : CommandBase
    {
        public CommandTutorial(Container container, Annotation annotation)
            : base(container, annotation, nameof(CommandTutorial))
        {
            // No specific execution conditions - tutorial should always be available
        }

        protected override void ExecuteCore(object value)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceName = "X21.x21-samples.xlsx";

            using (var stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream != null)
                {
                    var tempPath = Path.GetTempFileName();
                    tempPath = Path.ChangeExtension(tempPath, ".xlsx");

                    using (var fileStream = File.Create(tempPath))
                    {
                        stream.CopyTo(fileStream);
                    }

                    var excelApp = Container.Resolve<Application>();
                    excelApp.Workbooks.Open(tempPath);
                }
                else
                {
                    Logger.Info("Tutorial samples file not found as embedded resource.");
                }
            }
        }
    }
}
