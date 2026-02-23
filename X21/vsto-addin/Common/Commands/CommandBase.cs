using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Input;
using X21.Common.Data;
using X21.Common.Model;

namespace X21.Common.Commands
{
    public abstract class CommandBase : AnnotatedModelBase, ICommand
    {
        public string Id { get; }

        protected CommandBase(Container container, Annotation annotation, string id)
            : base(container)
        {
            Id = id;
            Annotation = annotation;
            ExecuteConditions = Enumerable.Empty<IExecuteCondition>();
        }

        public event EventHandler CanExecuteChanged;

        public virtual bool CanExecute(object parameter)
        {
            try
            {
                return CanExecuteCore(parameter);
            }
            catch (Exception ex)
            {
                Logging.Logger.LogException(ex);
                return false;
            }
        }
        public void Execute(object parameter)
        {
            try
            {
                ExecuteCore(parameter);
            }
            catch (Exception ex)
            {
                Logging.Logger.LogExceptionUI(ex);
            }
        }

        public virtual void NotifyCanExecuteChanged()
        {
            CanExecuteChanged?.Invoke(this, EventArgs.Empty);
        }

        public override string ToString()
        {
            return Annotation?.Title ?? base.ToString();
        }

        protected virtual bool CanExecuteCore(object value)
        {
            return !ExecuteConditions.Any(c => !c.CanExecute());
        }
        protected abstract void ExecuteCore(object value);

        protected IEnumerable<IExecuteCondition> ExecuteConditions { get; set; }
    }
}
