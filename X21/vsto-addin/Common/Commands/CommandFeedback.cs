using X21.Common.Commands;
using X21.Common.Data;
using X21.Common.Model;
using X21.Utils;
using System;
using System.Diagnostics;

namespace X21.Common.Commands
{
    /// <summary>
    /// Command to open the feedback website
    /// </summary>
    public class CommandFeedback : CommandBase
    {
        public CommandFeedback(Container container, Annotation annotation)
            : base(container, annotation, nameof(CommandFeedback))
        {
            // No specific execution conditions - feedback should always be available
        }

        protected override void ExecuteCore(object value)
        {
            var url = "https://feedback.kontext21.com";
            Process.Start(new ProcessStartInfo
            {
                FileName = url,
                UseShellExecute = true
            });
        }
    }
}
