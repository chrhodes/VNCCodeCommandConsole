using System;

using Prism.Commands;
using Prism.Events;

using VNC;
using VNC.Core.Mvvm;
using VNC.Core.Services;

using VNCCodeCommandConsole.Core.Events;

namespace CCC.CodeChecks.Presentation.ViewModels
{
    public class DesignChecksViewModel : EventViewModelBase, IDesignChecksViewModel, IInstanceCountVM
    {
        #region Constructors, Initialization, and Load

        public DesignChecksViewModel(
            IEventAggregator eventAggregator,
            IMessageDialogService messageDialogService) : base(eventAggregator, messageDialogService)
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            // TODO(crhodes)
            // Save constructor parameters here

            InitializeViewModel();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void InitializeViewModel()
        {
            Int64 startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            InstanceCountVM++;

            CallCodeCheckCommand = new DelegateCommand<string>(
                CallCodeCheck, CallCodeCheckCanExecute);

            Message = "DesignChecksViewModel says hello";

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Enums (none)


        #endregion

        #region Structures (none)


        #endregion

        #region Fields and Properties

        public DelegateCommand<string> CallCodeCheckCommand { get; private set; }

        private string _message;

        public string Message
        {
            get => _message;
            set
            {
                if (_message == value)
                    return;
                _message = value;
                OnPropertyChanged();
            }
        }

        // TODO(crhodes)
        // Switch this to an Enum
        private string _language = "CS";

        public string Language
        {
            get => _language;
            set
            {
                if (_language == value)
                    return;
                _language = value;
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

        #region Private Methods

        private bool CallCodeCheckCanExecute(string codeCheckMethod)
        {
            return true;
        }

        void CallCodeCheck(string codeCheckMethod)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            string metricClass = $"VNC.CodeAnalysis.DesignMetrics.{Language}.{codeCheckMethod},VNC.CodeAnalysis";

            EventAggregator.GetEvent<InvokeCodeCheckEvent>().Publish(metricClass);
            Message = metricClass;

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region IInstanceCount

        private static int _instanceCountVM;

        public int InstanceCountVM
        {
            get => _instanceCountVM;
            set => _instanceCountVM = value;
        }

        #endregion IInstanceCount
    }
}
