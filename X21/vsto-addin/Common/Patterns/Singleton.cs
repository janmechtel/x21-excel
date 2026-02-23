using System;

namespace X21.Common.Patterns
{
    public class Singleton<T> where T : class
    {
        public static T Instance
        {
            get
            {
                lock (_sync)
                {
                    if (ReferenceEquals(_sInstance, null))
                    {
                        _sInstance = (T)Activator.CreateInstance(typeof(T));
                    }

                    return _sInstance;
                }
            }
            set => _sInstance = value;
        }
        private static T _sInstance;
        private static object _sync = new object();
    }
}
