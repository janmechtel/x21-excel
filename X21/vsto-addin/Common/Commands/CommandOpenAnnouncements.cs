using X21.Common.Data;
using X21.Common.Model;
using System.Diagnostics;

namespace X21.Common.Commands
{
    /// <summary>
    /// Command to open the public announcements page.
    /// </summary>
    public class CommandOpenAnnouncements : CommandBase
    {
        private const string AnnouncementsUrl = "https://feedback.kontext21.com/announcements";

        public CommandOpenAnnouncements(Container container, Annotation annotation, string id)
            : base(container, annotation, id)
        {
        }

        protected override void ExecuteCore(object value)
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = AnnouncementsUrl,
                UseShellExecute = true
            });
        }
    }
}
