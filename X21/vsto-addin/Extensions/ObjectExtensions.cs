namespace X21.Extensions
{
    public static class ObjectExtensions
    {
        public static int SafeGetHashCode(this object obj)
        {
            return obj?.GetHashCode() ?? 0;
        }
    }
}
