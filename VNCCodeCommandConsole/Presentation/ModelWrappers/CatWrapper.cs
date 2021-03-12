using System;
using System.Collections.Generic;

using VNC.Core.Mvvm;

namespace VNCCodeCommandConsole.Presentation.ModelWrappers
{
    public class CatWrapper : ModelWrapper<Domain.Cat>
    {
        public CatWrapper(Domain.Cat model) : base(model)
        {
        }

        public int Id { get { return Model.Id; } }

        public string FieldString
        {
            get { return GetValue<string>(); }
            set { SetValue(value); }
        }

        public int? FieldInt
        {
            get { return GetValue<int?>(); }
            set { SetValue(value); }
        }

        public double? FieldDouble
        {
            get { return GetValue<double?>(); }
            set { SetValue(value); }
        }

        public Single? FieldSingle
        {
            get { return GetValue<Single?>(); }
            set { SetValue(value); }
        }

        public DateTime? FieldDate
        {
            get { return GetValue<DateTime?>(); }
            set { SetValue(value); }
        }

        public DateTime? FieldDate2
        {
            get { return GetValue<DateTime?>(); }
            set { SetValue(value); }
        }

        public Boolean? FieldBoolean
        {
            get { return GetValue<Boolean?>(); }
            set { SetValue(value); }
        }

        public int? FavoriteCatId
        {
            get { return GetValue<int?>(); }
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

        public byte[] RowVersion
        {
            get { return GetValue<byte[]>(); }
            set { SetValue(value); }
        }

        public int? FavoriteToyId
        {
            get { return GetValue<int?>(); }
            set { SetValue(value); }
        }

        protected override IEnumerable<string> ValidateProperty(string propertyName)
        {
            switch (propertyName)
            {
                case nameof(FieldString):
                    if (string.Equals(FieldString, "Robot", StringComparison.OrdinalIgnoreCase))
                    {
                        yield return "Robots are not what we expected!";
                    }
                    break;
            }
        }
    }
}
