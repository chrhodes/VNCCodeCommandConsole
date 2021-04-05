using System;
using System.Reflection;
using System.Text;

using Prism.Commands;
using Prism.Events;

using VNC;
using VNC.CodeAnalysis;
using VNC.Core.Mvvm;
using VNC.Core.Services;

using VNCCodeCommandConsole.Core.Events;

using static VNC.CodeAnalysis.Types;

using VNCCA = VNC.CodeAnalysis;
using VNCSW = VNC.CodeAnalysis.SyntaxWalkers;

namespace VNCCodeCommandConsole.Presentation.ViewModels
{
    public class SyntaxWalkerViewModel : EventViewModelBase, ISyntaxWalkerViewModel, IInstanceCountVM
    {

        #region Constructors, Initialization, and Load

        public SyntaxWalkerViewModel(
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

            SyntaxWalkerCommand = new DelegateCommand<string>(
                WalkerExecute, WalkerCanExecute);

            Message = "SyntaxWalkerViewModel says hello";

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

        private bool _headerIsCollapsed = false;

        public bool HeaderIsCollapsed
        {
            get => _headerIsCollapsed;
            set
            {
                if (_headerIsCollapsed == value)
                    return;
                _headerIsCollapsed = value;
                OnPropertyChanged();
            }
        }
        
        private WalkerPattern _syntaxNodeWalker = new WalkerPattern(
            controlHeader: "Find SyntaxNode Syntax",
            buttonContent: "SyntaxNode Walker",
            commandParameter: "SyntaxNodeWalker",
            regExLabel: "RegExNode");

        public WalkerPattern SyntaxNodeWalker
        {
            get => _syntaxNodeWalker;
            set
            {
                if (_syntaxNodeWalker == value)
                    return;
                _syntaxNodeWalker = value;
                OnPropertyChanged();
            }
        }

        private WalkerPattern _syntaxTokenWalker = new WalkerPattern(
            controlHeader: "Find SyntaxToken Syntax",
            buttonContent: "SyntaxToken Walker",
            commandParameter: "SyntaxTokenWalker",
            regExLabel: "RegExToken");

        public WalkerPattern SyntaxTokenWalker
        {
            get => _syntaxTokenWalker;
            set
            {
                if (_syntaxTokenWalker == value)
                    return;
                _syntaxTokenWalker = value;
                OnPropertyChanged();
            }
        }

        private WalkerPattern _syntaxTriviaWalker = new WalkerPattern(
            controlHeader: "Find SyntaxTrivia Syntax",
            buttonContent: "SyntaxTrivia Walker",
            commandParameter: "SyntaxTriviaWalker",
            regExLabel: "RegExTrivia");

        public WalkerPattern SyntaxTriviaWalker
        {
            get => _syntaxTriviaWalker;
            set
            {
                if (_syntaxTriviaWalker == value)
                    return;
                _syntaxTriviaWalker = value;
                OnPropertyChanged();
            }
        }

        private WalkerPattern _structuredTriviaWalker = new WalkerPattern(
            controlHeader: "Find StructuredTrivia Syntax",
            buttonContent: "StructuredTrivia Walker",
            commandParameter: "StructuredTriviaWalker",
            regExLabel: "RegEx");

        public WalkerPattern StructuredTriviaWalker
        {
            get => _structuredTriviaWalker;
            set
            {
                if (_structuredTriviaWalker == value)
                    return;
                _structuredTriviaWalker = value;
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

        public DelegateCommand<string> SyntaxWalkerCommand { get; set; }

        public void WalkerExecute(string walkerPropertyName)
        {
            Int64 startTicks = Log.EVENT("Enter", Common.LOG_CATEGORY);

            // NOTE(crhodes)
            // Use walkerPropertyName and Reflection on this class
            // to find the method to pass.  This allows us to have one method
            // to handle all the walkers that use the RegExSyntaxWalker

            // First get the CommandParameter property value
            // from the WalkerPattern property that corresponds to the walkerPropertyName

            PropertyInfo walkerPropertyInfo = this.GetType().GetProperty(walkerPropertyName);
            var walkerProperty = walkerPropertyInfo.GetValue(this);
            var commandParameter = ((WalkerPattern)walkerProperty).CommandParameter;

            Message = commandParameter;

            // Second use the commandParameter to find the appropriate Method
            // to pass as a delegate in the published event

            var methodName = $"Display{commandParameter}";

            // NOTE(crhodes)
            // This expects the method to be public.  Research how to find private methods
            MethodInfo displayWalkerMethod = this.GetType().GetMethod(methodName);
            SearchTreeCommand walkerMethodDelegate = (SearchTreeCommand)displayWalkerMethod.CreateDelegate(typeof(SearchTreeCommand), this);

            switch (Language)
            {
                case "CS":
                    EventAggregator.GetEvent<InvokeCSSyntaxWalkerEvent>().Publish(walkerMethodDelegate);
                    break;

                case "VB":
                    EventAggregator.GetEvent<InvokeVBSyntaxWalkerEvent>().Publish(walkerMethodDelegate);
                    break;

                default:
                    break;
            }

            Log.EVENT("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public bool WalkerCanExecute(string tag)
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        public StringBuilder DisplaySyntaxNodeWalker(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VNCSyntaxWalker(Microsoft.CodeAnalysis.SyntaxWalkerDepth.Node);

            commandConfiguration.CodeAnalysisOptions.SyntaxWalkerDepth = Microsoft.CodeAnalysis.SyntaxWalkerDepth.Node;
            commandConfiguration.WalkerPattern = SyntaxNodeWalker;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.Common.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplaySyntaxTokenWalker(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VNCSyntaxWalker(Microsoft.CodeAnalysis.SyntaxWalkerDepth.Token);

            commandConfiguration.CodeAnalysisOptions.SyntaxWalkerDepth = Microsoft.CodeAnalysis.SyntaxWalkerDepth.Token;
            commandConfiguration.WalkerPattern = SyntaxTokenWalker;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.Common.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplaySyntaxTriviaWalker(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VNCSyntaxWalker(Microsoft.CodeAnalysis.SyntaxWalkerDepth.Trivia);

            commandConfiguration.CodeAnalysisOptions.SyntaxWalkerDepth = Microsoft.CodeAnalysis.SyntaxWalkerDepth.Trivia;
            commandConfiguration.WalkerPattern = SyntaxTriviaWalker;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.Common.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayStructuredTriviaWalker(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VNCSyntaxWalker(Microsoft.CodeAnalysis.SyntaxWalkerDepth.StructuredTrivia);

            commandConfiguration.CodeAnalysisOptions.SyntaxWalkerDepth = Microsoft.CodeAnalysis.SyntaxWalkerDepth.StructuredTrivia;
            commandConfiguration.WalkerPattern = StructuredTriviaWalker;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.Common.InvokeVNCSyntaxWalker(walker, commandConfiguration);
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
