using System.Collections.Generic;
using X21.Common.Data;

namespace X21.Common.Model
{
    public class NotifiableObject<T> : AnnotatedModelBase
    {
        private IObjectStore<T> ObjectStore { get; set; }

        private NotifiableObject(Container container, IObjectStore<T> objectStore)
            : base(container)
        {
            ObjectStore = objectStore;
        }

        public NotifiableObject(Container container)
            : this(container, new ObjectStoreMemory<T>())
        {
        }

        public NotifiableObject(Container container, T value)
            : this(container)
        {
            Value = value;
        }

        public T Value
        {
            get => ObjectStore.Value;
            set
            {
                if (!EqualityComparer<T>.Default.Equals(ObjectStore.Value, value))
                {

                    ObjectStore.Value = value;
                    OnPropertyChanged();
                }
            }
        }
    }
}
