using System;

using VNC;
using VNC.Core.Mvvm;

namespace VNCCodeCommandConsole.Presentation.ViewModels
{
    public class RibbonShellViewModel : ViewModelBase
    {

        private string _title = "VNCCodeCommandConsole - RibbonShell";

        public string Title
        {
            get => _title;
            set
            {
                if (_title == value)
                    return;
                _title = value;
                OnPropertyChanged();
            }
        }

        public RibbonShellViewModel()
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        //public RibbonShellViewModel(I customerDataService)
        //{
        //    //_customerDataService = customerDataService;

        //    //Customers = new ObservableCollection<Customer>();
        //}
    }
}
