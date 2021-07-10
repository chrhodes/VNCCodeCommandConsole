using System;

using VNC;
using VNC.Core.Mvvm;

namespace VNCCodeCommandConsole.Presentation.ViewModels
{
    public class RibbonShellViewModel : ViewModelBase, IInstanceCountVM
    {

        #region Constructors, Initialization, and Load

        public RibbonShellViewModel()
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InitializeViewModel();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void InitializeViewModel()
        {
            Int64 startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            InstanceCountVM++;
            RuntimeVersion = Common.RuntimeVersion;

            // TODO(crhodes)
            //

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Enums (none)


        #endregion

        #region Structures (none)


        #endregion

        #region Fields and Properties (none)

        private string _runtimeVersion = ".NET Runtime Version";

        public string RuntimeVersion
        {
            get => _runtimeVersion;
            set
            {
                if (_runtimeVersion == value)
                    return;
                _runtimeVersion = value;
                OnPropertyChanged();
            }
        }

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

        #endregion

        #region Event Handlers (none)


        #endregion

        #region Public Methods (none)


        #endregion

        #region Protected Methods (none)


        #endregion

        #region Private Methods (none)


        #endregion

        #region IInstanceCount

        private static int _instanceCountVM;

        public int InstanceCountVM
        {
            get => _instanceCountVM;
            set => _instanceCountVM = value;
        }

        #endregion
    }
}
