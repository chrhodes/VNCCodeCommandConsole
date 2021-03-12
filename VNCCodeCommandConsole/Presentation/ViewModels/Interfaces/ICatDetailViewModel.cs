using System.Threading.Tasks;

using VNC.Core.Mvvm;

namespace VNCCodeCommandConsole.Presentation.ViewModels
{
    public interface ICatDetailViewModel : IViewModel
    {
        Task LoadAsync(int id);
    }
}
