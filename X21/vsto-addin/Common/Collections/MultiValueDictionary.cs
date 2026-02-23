using System.Collections.Generic;
using System.Linq;

namespace X21.Common.Collections
{
    public class MultiValueDictionary<TKey, TValue> : Dictionary<TKey, List<TValue>>
    {
        public void Add(TKey key, TValue value)
        {
            var values = default(List<TValue>);
            if (!TryGetValue(key, out values))
            {
                values = new List<TValue>();
                this.Add(key, values);
            }
            values.Add(value);
        }
        public bool ContainsValue(TValue value)
        {
            return this.Any(i => i.Value.Contains(value));
        }
        public IEnumerable<TValue> GetValues(TKey key)
        {
            var values = default(List<TValue>);
            if (TryGetValue(key, out values))
            {
                return values;
            }
            else
            {
                return Enumerable.Empty<TValue>();
            }
        }

        public IEnumerable<TValue> AllValues => Values.SelectMany(s => s);
    }
}
