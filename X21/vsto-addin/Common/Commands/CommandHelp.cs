using X21.Common.Commands;
using X21.Common.Data;
using X21.Common.Model;
using System.Diagnostics;

namespace X21.Common.Commands
{
    /// <summary>
    /// Command to open the online help manual.
    /// </summary>
    public class CommandHelp : CommandBase
    {
        private const string HelpUrl = "https://docs.kontext21.com";

        public CommandHelp(Container container, Annotation annotation)
            : base(container, annotation, nameof(CommandHelp))
        {
            // Help should always be available.
        }

        protected override void ExecuteCore(object value)
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = HelpUrl,
                UseShellExecute = true
            });
        }
    }
}
