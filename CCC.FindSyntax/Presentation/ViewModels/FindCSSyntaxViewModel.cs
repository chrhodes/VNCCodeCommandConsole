using System;
using System.Text;

using Prism.Commands;
using Prism.Events;

using VNC;
using VNC.CodeAnalysis;
using VNC.Core.Mvvm;
using VNC.Core.Services;

using VNCCodeCommandConsole.Core.Events;

using VNCCA = VNC.CodeAnalysis;
using VNCSW = VNC.CodeAnalysis.SyntaxWalkers;

namespace CCC.FindSyntax.Presentation.ViewModels
{
    public class FindCSSyntaxViewModel : EventViewModelBase, IInstanceCountVM
    {

        #region Constructors, Initialization, and Load

        public FindCSSyntaxViewModel(
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

            UsingDirectiveWalkerCommand = new DelegateCommand(
                UsingDirectiveWalkerExecute, UsingDirectiveWalkerCanExecute);

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Enums


        #endregion

        #region Structures


        #endregion

        #region Fields and Properties

        private bool _usingDirectiveUseRegEx;
        private string _usingDirectiveRegEx = ".*";

        public string UsingDirectiveRegEx
        {
            get => _usingDirectiveRegEx;
            set
            {
                if (_usingDirectiveRegEx == value)
                    return;
                _usingDirectiveRegEx = value;
                OnPropertyChanged();
            }
        }

        public bool UsingDirectiveUseRegEx
        {
            get => _usingDirectiveUseRegEx;
            set
            {
                if (_usingDirectiveUseRegEx == value)
                    return;
                _usingDirectiveUseRegEx = value;
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

        #region UsingStatementWalker

        public DelegateCommand UsingDirectiveWalkerCommand { get; set; }


        public void UsingDirectiveWalkerExecute()
        {
            Int64 startTicks = Log.EVENT("Enter", Common.LOG_CATEGORY);

            EventAggregator.GetEvent<InvokeCSSyntaxWalkerEvent>().Publish(DisplayUsingDirectiveWalkerCS);

            Log.EVENT("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public bool UsingDirectiveWalkerCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        private StringBuilder DisplayUsingDirectiveWalkerCS(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.CS.UsingDirective();

            commandConfiguration.UseRegEx = UsingDirectiveUseRegEx;
            commandConfiguration.RegEx = UsingDirectiveRegEx;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.CS.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        #endregion

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
