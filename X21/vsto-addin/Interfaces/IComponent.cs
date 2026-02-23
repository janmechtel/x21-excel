using X21.Common.Data;

namespace X21.Interfaces
{
    public interface IComponent
    {
        Container Container { get; }

        void Init();
        void Exit();
    }
}
