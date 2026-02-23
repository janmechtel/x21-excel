using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using X21.Common.Collections;
using X21.Common.Patterns;
using Microsoft.Office.Interop.Excel;

namespace X21.Common.Data
{
    public class ClassLoader : Singleton<ClassLoader>
    {
        public void Init(Assembly assembly)
        {
            if (ReferenceEquals(_classNameToType, null))
            {
                _classNameToType = new MultiValueDictionary<string, Type>();
            }

            foreach (var type in assembly.GetTypes().Where(IsTypeSupported))
            {
                Add(type);
            }
        }

        public void Add(Type type)
        {
            _classNameToType.Add(type.Name, type);

            foreach (var subtype in type.GetNestedTypes().Where(IsTypeSupported))
            {
                _classNameToType.Add(type.Name + "." + subtype.Name, subtype);
            }
        }
        public void Clear()
        {
            if (_classNameToType != null)
            {
                _classNameToType.Clear();
            }
        }

        private static bool IsTypeSupported(Type t)
        {
            return (t.IsClass || t.IsValueType) && !(t.IsAbstract && t.FullName.StartsWith("System")) && t.Name.First() != '<';
        }

        public T Create<T>(string name)
        {
            return (T)Activator.CreateInstance(TypeByString(name));
        }
        public T Create<T>(string name, params object[] args)
        {
            return (T)Activator.CreateInstance(TypeByString(name), args);
        }

        public object Create<T>(Type t, params object[] parameters)
        {
            if (t.IsInterface)
            {
                t = Types.First(t.IsAssignableFrom);
            }

            return Activator.CreateInstance(t, parameters);
        }



        public IEnumerable<Type> Types => _classNameToType.AllValues;

        public Type TypeByString(string name)
        {
            if (!_classNameToType.ContainsKey(name))
            {
                throw new ArgumentException($"'{name}' is an unknown classname");
            }
            return _classNameToType[name].Last();
        }

        private MultiValueDictionary<string, Type> _classNameToType;
    }
}
