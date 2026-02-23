using X21.Common.Commands;
using X21.Common.Data;
using X21.Common.Model;

namespace X21.TaskPane.Commands
{
    /// <summary>
    /// Command to toggle the chat taskpane visibility
    /// </summary>
    public class CommandToggleTaskPane : CommandBase
    {
        public CommandToggleTaskPane(Container container, Annotation annotation)
            : base(container, annotation, nameof(CommandToggleTaskPane))
        {
            // No specific execution conditions - taskpane toggle should always be available
        }

        public CommandToggleTaskPane(Container container, Annotation annotation, string customId)
            : base(container, annotation, customId)
        {
            // No specific execution conditions - taskpane toggle should always be available
        }

        protected override void ExecuteCore(object value)
        {
            var taskPaneManager = Container.Resolve<TaskPaneManager>();
            taskPaneManager.ToggleTaskPane();
        }
    }
}
