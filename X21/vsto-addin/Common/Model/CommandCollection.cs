using X21.Common.Commands;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace X21.Common.Model
{
    public class CommandCollection : ObservableCollection<CommandBase>
    {
        public CommandCollection()
        {
        }

        public CommandCollection(List<CommandBase> list)
            : base(list)
        {
        }

        public CommandBase this[string id] => this.FirstOrDefault(c => c.Id == id);
        public CommandBase this[Type type] => this.FirstOrDefault(c => c.GetType().Equals(type));

        public T CommandByType<T>() where T : CommandBase
        {
            return this.FirstOrDefault(c => typeof(T).IsAssignableFrom(c.GetType())) as T;
        }
    }
}
