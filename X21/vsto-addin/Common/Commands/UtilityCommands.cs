using X21.Common.Data;
using X21.Common.Model;
using X21.Interfaces;

namespace X21.Common.Commands
{
    /// <summary>
    /// Component that registers utility commands like folder opening commands
    /// </summary>
    public class UtilityCommands : Component, ICommands
    {
        public UtilityCommands(Container container) : base(container)
        {
        }

        protected override void InitCommands()
        {
            // Register Excel log folder command
            var excelLogFolderAnnotation = new Annotation(Container)
            {
                Title = "XLS Logs"
            };
            Commands.Add(new CommandOpenExcelLogFolder(Container, excelLogFolderAnnotation));

            // Register Deno log folder command
            var denoLogFolderAnnotation = new Annotation(Container)
            {
                Title = "Deno Logs"
            };
            Commands.Add(new CommandOpenDenoLogFolder(Container, denoLogFolderAnnotation));

            // Register binary folder command
            var binaryFolderAnnotation = new Annotation(Container)
            {
                Title = "Binary Folder"
            };
            Commands.Add(new CommandOpenBinaryFolder(Container, binaryFolderAnnotation));

            var slashCommandsAnnotation = new Annotation(Container)
            {
                Title = "Commands"
            };
            Commands.Add(new CommandCreateSlashCommandsSheet(Container, slashCommandsAnnotation));

            var localDataFolderAnnotation = new Annotation(Container)
            {
                Title = "Data Folder"
            };
            Commands.Add(new CommandOpenLocalDataFolder(Container, localDataFolderAnnotation));

            var helpAnnotation = new Annotation(Container)
            {
                Title = "Help"
            };
            Commands.Add(new CommandHelp(Container, helpAnnotation));

            // Register announcements command for the version split button
            var announcementsAnnotation = new Annotation(Container)
            {
                Title = "Announcements"
            };
            Commands.Add(new CommandOpenAnnouncements(Container, announcementsAnnotation, "VersionLabel"));

            // Register "What's New" command for dedicated button
            var whatsNewAnnotation = new Annotation(Container)
            {
                Title = "What's New?"
            };
            Commands.Add(new CommandOpenAnnouncements(Container, whatsNewAnnotation, "CommandWhatsNew"));
        }
    }
}
