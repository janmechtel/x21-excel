using X21.Common.Data;
using X21.Interfaces;

namespace X21.Common.Model
{
    public abstract class Component : ModelBase, IComponent, ICommands
    {
        public Component(Container container)
            : base(container)
        {
            Commands = new CommandCollection();
        }

        protected virtual void InitCommands()
        {
            // no default implementation
        }

        public CommandCollection Commands { get; }

        public override string ToString()
        {
            return GetType().Name;
        }

        public virtual void Init()
        {
            InitCommands();
        }

        public virtual void Exit()
        {
            // no default implementation
        }
    }
}
