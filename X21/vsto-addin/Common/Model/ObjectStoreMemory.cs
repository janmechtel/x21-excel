namespace X21.Common.Model
{
    public class ObjectStoreMemory<T> : IObjectStore<T>
    {
        public T Value { get; set; }
    }
}
