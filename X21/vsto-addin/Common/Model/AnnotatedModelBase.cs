using X21.Common.Data;

namespace X21.Common.Model
{
    public class AnnotatedModelBase : ModelBase
    {
        public AnnotatedModelBase(Container container)
            : base(container)
        {
        }

        public Annotation Annotation { get; set; }
    }
}
