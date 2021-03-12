using System;
using System.Collections.Generic;

using VNC.Core.Mvvm;

namespace VNCCodeCommandConsole.Presentation.ModelWrappers
{
    public class ToyWrapper : ModelWrapper<Domain.Toy>
    {
        public ToyWrapper(Domain.Toy model) : base(model)
        {
        }

        public int Id { get { return Model.Id; } }

        public string Name
        {
            get { return GetValue<string>(); }
            set { SetValue(value); }
        }

        public DateTime? DateCreated
        {
            get { return GetValue<DateTime?>(); }
            set { SetValue(value); }
        }

        public DateTime? DateModified
        {
            get { return GetValue<DateTime?>(); }
            set { SetValue(value); }
        }

        public Boolean? IsDirty
        {
            get { return GetValue<Boolean?>(); }
            set { SetValue(value); }
        }

        protected override IEnumerable<string> ValidateProperty(string propertyName)
        {
            switch (propertyName)
            {
                case nameof(Name):
                    if (string.Equals(Name, "Pickle", StringComparison.OrdinalIgnoreCase))
                    {
                        yield return "Pickles are not what we expected!";
                    }
                    break;
            }
        }
    }
}
