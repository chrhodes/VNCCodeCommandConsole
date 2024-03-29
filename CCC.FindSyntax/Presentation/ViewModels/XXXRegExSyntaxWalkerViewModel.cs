﻿using System;
using System.Text;

using Prism.Commands;
using Prism.Events;
using Prism.Services.Dialogs;

using VNC;
using VNC.CodeAnalysis;
using VNC.Core.Mvvm;
using VNC.Core.Services;

using VNCCodeCommandConsole.Core.Events;

using VNCCA = VNC.CodeAnalysis;
using VNCSW = VNC.CodeAnalysis.SyntaxWalkers;

namespace CCC.FindSyntax.Presentation.ViewModels
{
    public class XXXRegExSyntaxWalkerViewModel : EventViewModelBase, IInstanceCountVM
    {

        #region Constructors, Initialization, and Load

        public XXXRegExSyntaxWalkerViewModel(
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

            SyntaxWalkerCommand = new DelegateCommand(
                WalkerExecute, WalkerCanExecute);

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Enums


        #endregion

        #region Structures


        #endregion

        #region Fields and Properties

        private string _header;
        private CommandTypes.SearchTreeCommand _searchTreeCommand;

        private bool _useRegEx;
        private string _regEx = "RESW-VM.*";

        public string RegEx
        {
            get => _regEx;
            set
            {
                if (_regEx == value)
                    return;
                _regEx = value;
                OnPropertyChanged();
            }
        }

        public bool UseRegEx
        {
            get => _useRegEx;
            set
            {
                if (_useRegEx == value)
                    return;
                _useRegEx = value;
                OnPropertyChanged();
            }
        }
        private string _message = "RESW-VM-Initial Message";
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
        
        public string Header
        {
            get => _header;
            set
            {
                if (_header == value)
                    return;
                _header = value;
                OnPropertyChanged();
            }
        }

        public CommandTypes.SearchTreeCommand SearchTreeCommand
        {
            get => _searchTreeCommand;
            set
            {
                if (_searchTreeCommand == value)
                    return;
                _searchTreeCommand = value;
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

        #region ImportsStatementWalker Command

        public DelegateCommand SyntaxWalkerCommand { get; set; }

        public string WalkerContent { get; set; } = "SyntaxWalker";
        public string WalkerToolTip { get; set; } = "SyntaxWalker ToolTip";

        // Can get fancy and use Resources
        //public string ImportsStatementWalkerContent { get; set; } = "ViewName_ImportsStatementWalkerContent";
        //public string ImportsStatementWalkerToolTip { get; set; } = "ViewName_ImportsStatementWalkerContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_ImportsStatementWalkerContent">ImportsStatementWalker</system:String>
        //    <system:String x:Key="ViewName_ImportsStatementWalkerContentToolTip">ImportsStatementWalker ToolTip</system:String>  

        public void WalkerExecute()
        {
            Int64 startTicks = Log.EVENT("Enter", Common.LOG_CATEGORY);

            //Helper.ProcessOperation(DisplayImportsStatementWalkerVB, CodeExplorer, CodeExplorerContext, CodeExplorer.configurationOptions);

            Message = $"VM-{DateTime.Now.ToLongTimeString()}";

            //EventAggregator.GetEvent<InvokeVBSyntaxWalkerEvent>().Publish(SearchTreeCommand);

            //EventAggregator.GetEvent<SyntaxWalkerResultEvent>().Publish("Syntax Walker Result from ImportsStatementWalker");

            // If using events to tell something else to act

            //    EventAggregator.GetEvent<ImportsStatementWalkerEvent>().Publish(
            //        new ImportsStatementWalkerEventArgs()
            //        {
            //             Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //             Project = _contextMainViewModel.Context.SelectedProject,
            //             Team = _contextMainViewModel.Context.SelectedTeam
            //        });

            // Put this in places that listen for event and create method ImportsStatementWalker
            // EventAggregator.GetEvent<ImportsStatementWalkerEvent>().Subscribe(ImportsStatementWalker);

            Log.EVENT("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public bool WalkerCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
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
