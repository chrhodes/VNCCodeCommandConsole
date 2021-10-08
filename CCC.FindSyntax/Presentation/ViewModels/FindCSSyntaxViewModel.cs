using System;
using System.Reflection;
using System.Text;

using Prism.Commands;
using Prism.Events;
using Prism.Services.Dialogs;

using VNC;
using VNC.CodeAnalysis;
using VNC.CodeAnalysis.SyntaxWalkers;
using VNC.Core.Mvvm;

using VNCCodeCommandConsole.Core.Events;

using static VNC.CodeAnalysis.CommandTypes;

using VNCCA = VNC.CodeAnalysis;
using VNCSW = VNC.CodeAnalysis.SyntaxWalkers;

namespace CCC.FindSyntax.Presentation.ViewModels
{
    public class FindCSSyntaxViewModel : EventViewModelBase, IInstanceCountVM
    {

        #region Constructors, Initialization, and Load

        public FindCSSyntaxViewModel(
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

            SyntaxWalkerCommand = new DelegateCommand<string>(
                WalkerExecute, WalkerCanExecute);
            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Enums (none)


        #endregion

        #region Structures (none)


        #endregion

        #region Fields and Properties

        private WalkerPattern _usingDirectiveWalker = new WalkerPattern(
            controlHeader: "Find UsingDirective Syntax",
            buttonContent: "UsingDirective Walker",
            commandParameter: "UsingDirectiveWalker",
            regExLabel: "RegEx");

        public WalkerPattern UsingDirectiveWalker
        {
            get => _usingDirectiveWalker;
            set
            {
                if (_usingDirectiveWalker == value)
                    return;
                _usingDirectiveWalker = value;
                OnPropertyChanged();
            }
        }

        private WalkerPattern _namespaceDeclarationWalker = new WalkerPattern(
            controlHeader: "Find NamespaceDeclaration Syntax",
            buttonContent: "NamespaceDeclaration Walker",
            commandParameter: "NamespaceDeclarationWalker",
            regExLabel: "RegEx");

        public WalkerPattern NamespaceDeclarationWalker
        {
            get => _namespaceDeclarationWalker;
            set
            {
                if (_namespaceDeclarationWalker == value)
                    return;
                _namespaceDeclarationWalker = value;
                OnPropertyChanged();
            }
        }

        private WalkerPatternBlock _classDeclarationWalker = new WalkerPatternBlock(
            controlHeader: "Find ClassDeclaration Syntax",
            buttonContent: "ClassDeclaration Walker",
            commandParameter: "ClassDeclarationWalker",
            regExLabel: "RegEx",
            displayBlockLabel: "Display Class Block");

        public WalkerPatternBlock ClassDeclarationWalker
        {
            get => _classDeclarationWalker;
            set
            {
                if (_classDeclarationWalker == value)
                    return;
                _classDeclarationWalker = value;
                OnPropertyChanged();
            }
        }

        private WalkerPatternBlock _methodDeclarationWalker = new WalkerPatternBlock(
            controlHeader: "Find MethodDeclaration Syntax",
            buttonContent: "MethodDeclaration Walker",
            commandParameter: "MethodDeclarationWalker",
            regExLabel: "RegEx",
            displayBlockLabel: "Display Method Block");

        public WalkerPatternBlock MethodDeclarationWalker
        {
            get => _methodDeclarationWalker;
            set
            {
                if (_methodDeclarationWalker == value)
                    return;
                _methodDeclarationWalker = value;
                OnPropertyChanged();
            }
        }

        private WalkerPattern _fieldDeclarationWalker = new WalkerPattern(
            controlHeader: "Find FieldDeclaration Syntax",
            buttonContent: "FieldDeclaration Walker",
            commandParameter: "FieldDeclarationWalker",
            regExLabel: "RegEx");

        public WalkerPattern FieldDeclarationWalker
        {
            get => _fieldDeclarationWalker;
            set
            {
                if (_fieldDeclarationWalker == value)
                    return;
                _fieldDeclarationWalker = value;
                OnPropertyChanged();
            }
        }

        private WalkerPatternBlock _propertyDeclarationWalker = new WalkerPatternBlock(
            controlHeader: "Find PropertyDeclaration Syntax",
            buttonContent: "PropertyDeclaration Walker",
            commandParameter: "PropertyDeclarationWalker",
            regExLabel: "RegEx",
            displayBlockLabel: "Display Property Block");

        public WalkerPatternBlock PropertyDeclarationWalker
        {
            get => _propertyDeclarationWalker;
            set
            {
                if (_propertyDeclarationWalker == value)
                    return;
                _propertyDeclarationWalker = value;
                OnPropertyChanged();
            }
        }

        private WalkerPatternStruct _structDeclarationWalker = new WalkerPatternStruct(
            controlHeader: "Find StructDeclaration Syntax",
            buttonContent: "StructDeclaration Walker",
            commandParameter: "StructDeclarationWalker",
            regExLabel: "RegEx");

        public WalkerPatternStruct StructDeclarationWalker
        {
            get => _structDeclarationWalker;
            set
            {
                if (_structDeclarationWalker == value)
                    return;
                _structDeclarationWalker = value;
                OnPropertyChanged();
            }
        }

        private WalkerPattern _localDeclarationStatementWalker = new WalkerPattern(
            controlHeader: "Find LocalDeclarationStatement Syntax",
            buttonContent: "LocalDeclarationStatement Walker",
            commandParameter: "LocalDeclarationStatementWalker",
            regExLabel: "RegEx");

        public WalkerPattern LocalDeclarationStatementWalker
        {
            get => _localDeclarationStatementWalker;
            set
            {
                if (_localDeclarationStatementWalker == value)
                    return;
                _localDeclarationStatementWalker = value;
                OnPropertyChanged();
            }
        }


    private WalkerPatternBlock _ifStatementWalker = new WalkerPatternBlock(
        controlHeader: "Find IfStatement Syntax",
        buttonContent: "IfStatement Walker",
        commandParameter: "IfStatementWalker",
        regExLabel: "RegEx",
        displayBlockLabel: "Display IfStatement Block");

        public WalkerPatternBlock IfStatementWalker
        {
            get => _ifStatementWalker;
            set
            {
                if (_ifStatementWalker == value)
                    return;
                _ifStatementWalker = value;
                OnPropertyChanged();
            }
        }

        private bool _headerIsCollapsed = true;

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

        #endregion

        #region Event Handlers (none)


        #endregion

        #region Public Methods (none)


        #endregion

        #region Protected Methods (none)


        #endregion

        #region Private Methods

        public DelegateCommand<string> SyntaxWalkerCommand { get; set; }

        public void WalkerExecute(string walkerPropertyName)
        {
            Int64 startTicks = Log.EVENT("Enter", Common.LOG_CATEGORY);

            // NOTE(crhodes)
            // Use walkerPropertyName and Reflection on this class
            // to find the method to pass.
            // This allows us to have one method
            // to handle all the walkers that use the RegExSyntaxWalker

            // First get the CommandParameter property value
            // from the WalkerPattern property that corresponds to the walkerPropertyName

            PropertyInfo walkerPropertyInfo = this.GetType().GetProperty(walkerPropertyName);
            var walkerProperty = walkerPropertyInfo.GetValue(this);
            var commandParameter = ((WalkerPattern)walkerProperty).CommandParameter;

            Message = commandParameter;

            // Second use the commandParameter to find the appropriate Method
            // to pass as a delegate in the published event

            var methodName = $"Display{commandParameter}CS";

            // NOTE(crhodes)
            // This expects the method to be public.  Research how to find private methods
            MethodInfo displayWalkerMethod = this.GetType().GetMethod(methodName);
            SearchTreeCommand walkerMethodDelegate = (SearchTreeCommand)displayWalkerMethod.CreateDelegate(typeof(SearchTreeCommand), this);
            EventAggregator.GetEvent<InvokeCSSyntaxWalkerEvent>().Publish(walkerMethodDelegate);

            Log.EVENT("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public bool WalkerCanExecute(string tag)
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        public StringBuilder DisplayUsingDirectiveWalkerCS(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.CS.UsingDirective();
            
            commandConfiguration.WalkerPattern = UsingDirectiveWalker;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.CS.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayNamespaceDeclarationWalkerCS(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.CS.NamespaceDeclaration();

            commandConfiguration.WalkerPattern = NamespaceDeclarationWalker;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.CS.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayClassDeclarationWalkerCS(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            VNCSW.CS.VNCCSSyntaxWalkerBase walker = new VNCSW.CS.ClassDeclaration();

            commandConfiguration.WalkerPattern = ClassDeclarationWalker;
            commandConfiguration.CodeAnalysisOptions.DisplayStatementBlock = ClassDeclarationWalker.DisplayBlock;

            // TODO(crhodes)
            // Figure this out Does this matter in CS?

            //if (ClassDeclarationWalker.DisplayBlock)
            //{
            //    walker = new VNCSW.CS.ClassBlock();
            //}
            //else
            //{
            //    walker = new VNCSW.CS.ClassStatement();
            //}

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.CS.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayMethodDeclarationWalkerCS(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            VNCSW.CS.VNCCSSyntaxWalkerBase walker = new VNCSW.CS.MethodDeclaration();

            commandConfiguration.WalkerPattern = MethodDeclarationWalker;
            commandConfiguration.CodeAnalysisOptions.DisplayStatementBlock = MethodDeclarationWalker.DisplayBlock;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.CS.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayFieldDeclarationWalkerCS(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            // TODO(crhodes)
            // Update this from UI selections

            VNCSW.CS.VNCCSSyntaxWalkerBase walker = new VNCSW.CS.FieldDeclaration(VNCCA.SyntaxNode.FieldDeclarationLocation.Class);

            commandConfiguration.WalkerPattern = FieldDeclarationWalker;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.CS.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayPropertyDeclarationWalkerCS(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            VNCSW.CS.VNCCSSyntaxWalkerBase walker = new VNCSW.CS.PropertyDeclaration();

            commandConfiguration.WalkerPattern = PropertyDeclarationWalker;
            commandConfiguration.CodeAnalysisOptions.DisplayStatementBlock = PropertyDeclarationWalker.DisplayBlock;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.CS.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayStructDeclarationWalkerCS(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            VNCSW.CS.VNCCSSyntaxWalkerBase walker = new VNCSW.CS.StructDeclaration();

            commandConfiguration.WalkerPattern = StructDeclarationWalker;
            commandConfiguration.CodeAnalysisOptions.DisplayFields = StructDeclarationWalker.DisplayFields;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.CS.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayIfStatementWalkerCS(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            VNCSW.CS.VNCCSSyntaxWalkerBase walker = new VNCSW.CS.IfStatement();

            commandConfiguration.WalkerPattern = IfStatementWalker;
            commandConfiguration.CodeAnalysisOptions.DisplayStatementBlock = IfStatementWalker.DisplayBlock;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.CS.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayLocalDeclarationStatementWalkerCS(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            VNCSW.CS.VNCCSSyntaxWalkerBase walker = new VNCSW.CS.LocalDeclarationStatement();

            commandConfiguration.WalkerPattern = LocalDeclarationStatementWalker;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.CS.InvokeVNCSyntaxWalker(walker, commandConfiguration);
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
