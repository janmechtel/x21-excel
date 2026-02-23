using X21.Common.Commands;
using X21.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;

namespace X21.Common.Data
{
    public class Container
    {
        private readonly Dictionary<Type, object> _registrations = new Dictionary<Type, object>();

        public virtual CommandBase CommandById(string id)
        {

            var commands = ResolveAll<ICommands>()
                .SelectMany(c => c.Commands)
                .Where(c => c.Id == id);

            return commands.FirstOrDefault();
        }

        public void RegisterSingleton<T>(T obj)
        {
            _registrations[typeof(T)] = obj;
        }

        public void RegisterSingleton<T>() where T : IComponent
        {
            var instance = CreateByType<T>(this);
            _registrations[typeof(T)] = instance;
        }

        public T Resolve<T>()
        {
            if (_registrations.TryGetValue(typeof(T), out var factory))
            {
                return (T)factory;
            }

            throw new InvalidOperationException($"Service of type {typeof(T)} is not registered.");
        }

        public IEnumerable<T> ResolveAll<T>()
        {
            foreach (var registration in _registrations)
            {
                if (typeof(T).IsAssignableFrom(registration.Key))
                {
                    yield return (T)registration.Value;
                }
            }
        }

        public T CreateByType<T>(params object[] parameters)
        {
            return (T)ClassLoader.Instance.Create<T>(typeof(T), parameters);
        }
    }
}
