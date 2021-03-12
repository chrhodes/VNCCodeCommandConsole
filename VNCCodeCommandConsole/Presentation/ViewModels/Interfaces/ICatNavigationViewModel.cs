using System.Threading.Tasks;

using VNC.Core.Mvvm;

namespace VNCCodeCommandConsole.Presentation.ViewModels
{
    public interface ICatNavigationViewModel : IViewModel
    {
        Task LoadAsync();
    }
}
