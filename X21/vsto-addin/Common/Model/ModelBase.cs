using System.ComponentModel;
using System.Runtime.CompilerServices;
using X21.Logging;
using Container = X21.Common.Data.Container;

namespace X21.Common.Model
{
    public class ModelBase : INotifyPropertyChanged
    {
        public ModelBase(Container container)
        {
            Container = container;
        }

        public Container Container { get; set; }

        public virtual ILogger Logger => Logging.Logger.Instance;

        public event PropertyChangedEventHandler PropertyChanged;
        public event PropertyChangedEventHandler PropertyChangedInit
        {
            add
            {
                value(this, null);
                PropertyChanged += value;
            }
            remove
            {
                PropertyChanged -= value;
            }
        }

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
