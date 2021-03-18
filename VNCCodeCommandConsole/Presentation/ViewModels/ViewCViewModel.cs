using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Input;

using Prism.Commands;
using Prism.Events;

using VNC;
using VNC.Core.Events;
using VNC.Core.Mvvm;
using VNC.Core.Services;

namespace VNCCodeCommandConsole.Presentation.ViewModels
{
    public class ViewCViewModel : EventViewModelBase, IViewABCViewModel, IInstanceCountVM
    {

        #region Constructors, Initialization, and Load

        public ViewCViewModel(
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

            // TODO(crhodes)
            //

            Message = "ViewModelC says hello";
            MessageA = "ViewModelC says hello A";
            MessageB = "ViewModelC says hello B";
            MessageC = "ViewModelC says hello C";

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Enums


        #endregion

        #region Structures


        #endregion

        #region Fields and Properties

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

        private string _messageA;

        public string MessageA
        {
            get => _messageA;
            set
            {
                if (_messageA == value)
                    return;
                _messageA = value;
                OnPropertyChanged();
            }
        }

        private string _messageB;

        public string MessageB
        {
            get => _messageB;
            set
            {
                if (_messageB == value)
                    return;
                _messageB = value;
                OnPropertyChanged();
            }
        }

        private string _messageC;

        public string MessageC
        {
            get => _messageC;
            set
            {
                if (_messageC == value)
                    return;
                _messageC = value;
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
