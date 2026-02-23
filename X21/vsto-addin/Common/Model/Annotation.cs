using System;
using X21.Common.Data;
using X21.Extensions;

namespace X21.Common.Model
{
    public class Annotation : ModelBase, IEquatable<Annotation>
    {
        public Annotation(Container container)
            : base(container)
        {

        }

        public string Title
        {
            get => _title;
            set
            {
                if (_title != value)
                {
                    _title = value;
                    OnPropertyChanged();
                }
            }
        }
        private string _title;

        public string Description
        {
            get => _description;
            set
            {
                if (_description != value)
                {
                    _description = value;
                    OnPropertyChanged();
                }
            }
        }
        private string _description;

        public string ToolTip
        {
            get => _toolTip;
            set
            {
                if (_toolTip != value)
                {
                    _toolTip = value;
                    OnPropertyChanged();
                }
            }
        }
        private string _toolTip;

        public override bool Equals(object obj)
        {
            return Equals(obj as Annotation);
        }
        public bool Equals(Annotation other)
        {
            return other != null
                   && Title.EqualsIgnoreCaseSafe(other.Title)
                   && Description.EqualsIgnoreCaseSafe(other.Description)
                   && ToolTip.EqualsIgnoreCaseSafe(other.ToolTip);
        }
        public override int GetHashCode()
        {
            return Title.SafeGetHashCode()
                   ^ Description.SafeGetHashCode()
                   ^ ToolTip.SafeGetHashCode();
        }

        public static Annotation Empty(Container container) => new Annotation(container);
    }
}
