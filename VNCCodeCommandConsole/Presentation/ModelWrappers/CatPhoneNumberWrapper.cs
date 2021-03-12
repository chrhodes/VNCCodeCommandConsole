using VNC.Core.Mvvm;

using VNCCodeCommandConsole.Domain;

namespace VNCCodeCommandConsole.Presentation.ModelWrappers
{
    public class CatPhoneNumberWrapper : ModelWrapper<CatPhoneNumber>
    {
        public CatPhoneNumberWrapper(CatPhoneNumber model) : base(model)
        {
        }

        public string Number
        {
            get { return GetValue<string>(); }
            set { SetValue(value); }
        }

    }
}
