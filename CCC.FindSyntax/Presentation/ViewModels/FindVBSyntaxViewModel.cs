using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

using Prism.Commands;
using Prism.Events;

using VNC;
using VNC.Core.Events;
using VNC.Core.Mvvm;
using VNC.Core.Services;

using VNCCodeCommandConsole.Core.Events;

using VNCCA = VNC.CodeAnalysis;
using VNCSW = VNC.CodeAnalysis.SyntaxWalkers;

namespace CCC.FindSyntax.Presentation.ViewModels
{
    public class FindVBSyntaxViewModel : EventViewModelBase, IInstanceCountVM
    {

        #region Constructors, Initialization, and Load

        public FindVBSyntaxViewModel(
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

            SayHelloCommand = new DelegateCommand(
                SayHello, SayHelloCanExecute);

            ImportsStatementWalkerCommand = new DelegateCommand(
                ImportsStatementWalkerExecute, ImportsStatementWalkerCanExecute);

            Message = "FindVBSyntaxViewModel says hello";

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Enums


        #endregion

        #region Structures


        #endregion

        #region Fields and Properties

        public ICommand SayHelloCommand { get; private set; }

        private bool _importsStatementUseRegEx;
        private string _importsStatementRegEx = ".*";
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


        public string ImportsStatementRegEx
        {
            get => _importsStatementRegEx;
            set
            {
                if (_importsStatementRegEx == value)
                    return;
                _importsStatementRegEx = value;
                OnPropertyChanged();
            }
        }

        public bool ImportsStatementUseRegEx
        {
            get => _importsStatementUseRegEx;
            set
            {
                if (_importsStatementUseRegEx == value)
                    return;
                _importsStatementUseRegEx = value;
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

        public DelegateCommand ImportsStatementWalkerCommand { get; set; }
        public string ImportsStatementWalkerContent { get; set; } = "ImportsStatementWalker";
        public string ImportsStatementWalkerToolTip { get; set; } = "ImportsStatementWalker ToolTip";

        // Can get fancy and use Resources
        //public string ImportsStatementWalkerContent { get; set; } = "ViewName_ImportsStatementWalkerContent";
        //public string ImportsStatementWalkerToolTip { get; set; } = "ViewName_ImportsStatementWalkerContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_ImportsStatementWalkerContent">ImportsStatementWalker</system:String>
        //    <system:String x:Key="ViewName_ImportsStatementWalkerContentToolTip">ImportsStatementWalker ToolTip</system:String>  

        public void ImportsStatementWalkerExecute()
        {
            Int64 startTicks = Log.EVENT("Enter", Common.LOG_CATEGORY);
            // TODO(crhodes)
            // Do something amazing.
            Message = $"ImportsStatementWalkerExecute UseRegEx({ImportsStatementUseRegEx}) RegEx({ImportsStatementRegEx})";

            //Helper.ProcessOperation(DisplayImportsStatementWalkerVB, CodeExplorer, CodeExplorerContext, CodeExplorer.configurationOptions);

            EventAggregator.GetEvent<InvokeSyntaxWalkerEvent>().Publish(DisplayImportsStatementWalkerVB);

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

        public bool ImportsStatementWalkerCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }


        #endregion

        private bool SayHelloCanExecute()
        {
            return true;
        }

        void SayHello()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Message = "Hello";

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private StringBuilder DisplayImportsStatementWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.ImportsStatement();

            commandConfiguration.UseRegEx = ImportsStatementUseRegEx;
            commandConfiguration.RegEx = ImportsStatementRegEx;
            // TODO(crhodes)
            // This needs to come in with the commandConfiguration

            //commandConfiguration.ConfigurationOptions = CodeExplorer.configurationOptions.GetConfigurationInfo();

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

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
