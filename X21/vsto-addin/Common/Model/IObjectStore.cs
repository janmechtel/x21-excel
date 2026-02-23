namespace X21.Common.Model
{
    public interface IObjectStore<T>
    {
        T Value { get; set; }
    }
}
