using System;

using Prism.Commands;
using Prism.Events;
using Prism.Services.Dialogs;

using VNC;
using VNC.Core.Mvvm;
using VNC.Core.Services;

using VNCCodeCommandConsole.Core.Events;

namespace CCC.CodeChecks.Presentation.ViewModels
{
    public class QualityChecksViewModel : EventViewModelBase, IQualityChecksViewModel, IInstanceCountVM
    {

        #region Constructors, Initialization, and Load

        public QualityChecksViewModel(
            IEventAggregator eventAggregator,
            IDialogService dialogService) : base(eventAggregator, dialogService)
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

            // TODO(crhodes)
            //

            CallCodeCheckCommand = new DelegateCommand<string>(
                CallCodeCheck, CallCodeCheckCanExecute);

            Message = "QualityChecksViewModel says hello";

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

        #region Event Handlers


        #endregion

        #region Public Methods


        #endregion

        #region Protected Methods


        #endregion

        #region Private Methods

        private bool CallCodeCheckCanExecute(string codeCheckMethod)
        {
            return true;
        }

        void CallCodeCheck(string codeCheckMethod)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            string language = Language;

            // TODO(crhodes)
            // Update UI to bind to Language.

            string metricClass = $"VNC.CodeAnalysis.QualityMetrics.{language}.{codeCheckMethod},VNC.CodeAnalysis";

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
