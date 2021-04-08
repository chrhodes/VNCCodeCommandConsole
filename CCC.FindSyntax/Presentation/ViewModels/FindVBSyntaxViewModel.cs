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

        // This does not work.  Needs to be a property.
        //WalkerPattern ImportsStatementWalker = new WalkerPattern();

        private bool _checkIfDone;
        private bool _checkIf;
        private WalkerPattern _importsStatementWalker = new WalkerPattern(
            controlHeader: "Find ImportsStatement Syntax",
            buttonContent: "ImportsStatement Walker",
            commandParameter: "ImportsStatementWalker",
            regExLabel: "Imports");

        public WalkerPattern ImportsStatementWalker
        {
            get => _importsStatementWalker;
            set
            {
                if (_importsStatementWalker == value)
                    return;
                _importsStatementWalker = value;
                OnPropertyChanged();
            }
        }

        private WalkerPattern _namespaceStatementWalker = new WalkerPattern(
            controlHeader: "Find NamespaceStatement Syntax",
            buttonContent: "NamespaceStatement Walker",
            commandParameter: "NamespaceStatementWalker",
            regExLabel: "Namespace");

        public WalkerPattern NamespaceStatementWalker
        {
            get => _namespaceStatementWalker;
            set
            {
                if (_namespaceStatementWalker == value)
                    return;
                _namespaceStatementWalker = value;
                OnPropertyChanged();
            }
        }

        private BlockWalkerPattern _classStatementWalker = new BlockWalkerPattern(
            controlHeader: "Find ClassStatement Syntax",
            buttonContent: "ClassStatement Walker",
            commandParameter: "ClassStatementWalker",
            regExLabel: "Class",
            displayBlockLabel: "Display Class Block");

        public BlockWalkerPattern ClassStatementWalker
        {
            get => _classStatementWalker;
            set
            {
                if (_classStatementWalker == value)
                    return;
                _classStatementWalker = value;
                OnPropertyChanged();
            }
        }

        private BlockWalkerPattern _moduleStatementWalker = new BlockWalkerPattern(
            controlHeader: "Find ModuleStatement Syntax",
            buttonContent: "ModuleStatement Walker",
            commandParameter: "ModuleStatementWalker",
            regExLabel: "Module",
            displayBlockLabel: "Display Module Block");

        public BlockWalkerPattern ModuleStatementWalker
        {
            get => _moduleStatementWalker;
            set
            {
                if (_moduleStatementWalker == value)
                    return;
                _moduleStatementWalker = value;
                OnPropertyChanged();
            }
        }

        private BlockWalkerPattern _methodStatementWalker = new BlockWalkerPattern(
            controlHeader: "Find MethodStatement Syntax",
            buttonContent: "MethodStatement Walker",
            commandParameter: "MethodStatementWalker",
            regExLabel: "Method",
            displayBlockLabel: "Display Method Block");

        public BlockWalkerPattern MethodStatementWalker
        {
            get => _methodStatementWalker;
            set
            {
                if (_methodStatementWalker == value)
                    return;
                _methodStatementWalker = value;
                OnPropertyChanged();
            }
        }

        private BlockWalkerPattern _propertyStatementWalker = new BlockWalkerPattern(
            controlHeader: "Find PropertyStatement Syntax",
            buttonContent: "PropertyStatement Walker",
            commandParameter: "PropertyStatementWalker",
            regExLabel: "Property",
            displayBlockLabel: "Display Property Block");

        public BlockWalkerPattern PropertyStatementWalker
        {
            get => _propertyStatementWalker;
            set
            {
                if (_propertyStatementWalker == value)
                    return;
                _propertyStatementWalker = value;
                OnPropertyChanged();
            }
        }

        private WalkerPattern _handlesClauseWalker = new WalkerPattern(
            controlHeader: "Find HandlesClause Syntax",
            buttonContent: "HandlesClause Walker",
            commandParameter: "HandlesClauseWalker",
            regExLabel: "Handles");

        public WalkerPattern HandlesClauseWalker
        {
            get => _handlesClauseWalker;
            set
            {
                if (_handlesClauseWalker == value)
                    return;
                _handlesClauseWalker = value;
                OnPropertyChanged();
            }
        }

        private WalkerPattern _parameterListWalker = new WalkerPattern(
            controlHeader: "Find ParameterList Syntax",
            buttonContent: "ParameterList Walker",
            commandParameter: "ParameterListWalker",
            regExLabel: "Parameters");

        public WalkerPattern ParameterListWalker
        {
            get => _parameterListWalker;
            set
            {
                if (_parameterListWalker == value)
                    return;
                _parameterListWalker = value;
                OnPropertyChanged();
            }
        }

        private WalkerPattern _localDeclarationStatementWalker = new WalkerPattern(
            controlHeader: "Find LocalDeclarationStatement Syntax",
            buttonContent: "LocalDeclarationStatement Walker",
            commandParameter: "LocalDeclarationStatementWalker",
            regExLabel: "Declaration");

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

        private WalkerPattern _multiLineLambdaExpressionWalker = new WalkerPattern(
            controlHeader: "Find MultiLineLambdaExpression Syntax",
            buttonContent: "MultiLineLambdaExpression Walker",
            commandParameter: "MultiLineLambdaExpressionWalker",
            regExLabel: "RegEx");

        public WalkerPattern MultiLineLambdaExpressionWalker
        {
            get => _multiLineLambdaExpressionWalker;
            set
            {
                if (_multiLineLambdaExpressionWalker == value)
                    return;
                _multiLineLambdaExpressionWalker = value;
                OnPropertyChanged();
            }
        }

        private WalkerPattern _singleLineLambdaExpressionWalker = new WalkerPattern(
            controlHeader: "Find SingleLineLambdaExpression Syntax",
            buttonContent: "SingleLineLambdaExpression Walker",
            commandParameter: "SingleLineLambdaExpressionWalker",
            regExLabel: "RegEx");

        public WalkerPattern SingleLineLambdaExpressionWalker
        {
            get => _singleLineLambdaExpressionWalker;
            set
            {
                if (_singleLineLambdaExpressionWalker == value)
                    return;
                _singleLineLambdaExpressionWalker = value;
                OnPropertyChanged();
            }
        }

        private WalkerPattern _expressionStatementWalker = new WalkerPattern(
            controlHeader: "Find ExpressionStatement Syntax",
            buttonContent: "ExpressionStatement Walker",
            commandParameter: "ExpressionStatementWalker",
            regExLabel: "RegEx");

        public WalkerPattern ExpressionStatementWalker
        {
            get => _expressionStatementWalker;
            set
            {
                if (_expressionStatementWalker == value)
                    return;
                _expressionStatementWalker = value;
                OnPropertyChanged();
            }
        }

        private WalkerPattern _invocationExpressionWalker = new WalkerPattern(
            controlHeader: "Find InvocationExpression Syntax",
            buttonContent: "InvocationExpression Walker",
            commandParameter: "InvocationExpressionWalker",
            regExLabel: "RegEx");

        public WalkerPattern InvocationExpressionWalker
        {
            get => _invocationExpressionWalker;
            set
            {
                if (_invocationExpressionWalker == value)
                    return;
                _invocationExpressionWalker = value;
                OnPropertyChanged();
            }
        }

        private WalkerPattern _memberAccessExpressionWalker = new WalkerPattern(
            controlHeader: "Find MemberAccessExpression Syntax",
            buttonContent: "MemberAccessExpression Walker",
            commandParameter: "MemberAccessExpressionWalker",
            regExLabel: "RegEx");

        public WalkerPattern MemberAccessExpressionWalker
        {
            get => _memberAccessExpressionWalker;
            set
            {
                if (_memberAccessExpressionWalker == value)
                    return;
                _memberAccessExpressionWalker = value;
                OnPropertyChanged();
            }
        }

        private WalkerPattern _argumentListWalker = new WalkerPattern(
            controlHeader: "Find ArgumentList Syntax",
            buttonContent: "ArgumentList Walker",
            commandParameter: "ArgumentListWalker",
            regExLabel: "RegEx");

        public WalkerPattern ArgumentListWalker
        {
            get => _argumentListWalker;
            set
            {
                if (_argumentListWalker == value)
                    return;
                _argumentListWalker = value;
                OnPropertyChanged();
            }
        }

        private WalkerPattern _variableDeclaratorWalker = new WalkerPattern(
            controlHeader: "Find VariableDeclarator Syntax",
            buttonContent: "VariableDeclarator Walker",
            commandParameter: "VariableDeclaratorWalker",
            regExLabel: "RegEx");

        public WalkerPattern VariableDeclaratorWalker
        {
            get => _variableDeclaratorWalker;
            set
            {
                if (_variableDeclaratorWalker == value)
                    return;
                _variableDeclaratorWalker = value;
                OnPropertyChanged();
            }
        }

        private WalkerPattern _objectCreationExpressionWalker = new WalkerPattern(
            controlHeader: "Find ObjectCreationExpression Syntax",
            buttonContent: "ObjectCreationExpression Walker",
            commandParameter: "ObjectCreationExpressionWalker",
            regExLabel: "RegEx");

        public WalkerPattern ObjectCreationExpressionWalker
        {
            get => _objectCreationExpressionWalker;
            set
            {
                if (_objectCreationExpressionWalker == value)
                    return;
                _objectCreationExpressionWalker = value;
                OnPropertyChanged();
            }
        }

        private WalkerPattern _binaryExpressionWalker = new WalkerPattern(
            controlHeader: "Find BinaryExpression Syntax",
            buttonContent: "BinaryExpression Walker",
            commandParameter: "BinaryExpressionWalker",
            regExLabel: "RegEx");

        public WalkerPattern BinaryExpressionWalker
        {
            get => _binaryExpressionWalker;
            set
            {
                if (_binaryExpressionWalker == value)
                    return;
                _binaryExpressionWalker = value;
                OnPropertyChanged();
            }
        }

        private WalkerPattern _asNewClauseWalker = new WalkerPattern(
            controlHeader: "Find AsNewClause Syntax",
            buttonContent: "AsNewClause Walker",
            commandParameter: "AsNewClauseWalker",
            regExLabel: "RegEx");

        public WalkerPattern AsNewClauseWalker
        {
            get => _asNewClauseWalker;
            set
            {
                if (_asNewClauseWalker == value)
                    return;
                _asNewClauseWalker = value;
                OnPropertyChanged();
            }
        }

        private WalkerPattern _simpleAsClauseWalker = new WalkerPattern(
            controlHeader: "Find SimpleAsClause Syntax",
            buttonContent: "SimpleAsClause Walker",
            commandParameter: "SimpleAsClauseWalker",
            regExLabel: "RegEx");

        public WalkerPattern SimpleAsClauseWalker
        {
            get => _simpleAsClauseWalker;
            set
            {
                if (_simpleAsClauseWalker == value)
                    return;
                _simpleAsClauseWalker = value;
                OnPropertyChanged();
            }
        }

        private WalkerPattern _syntaxNodeWalker = new WalkerPattern(
            controlHeader: "Find SyntaxNode Syntax",
            buttonContent: "SyntaxNode Walker",
            commandParameter: "SyntaxNodeWalker",
            regExLabel: "RegEx");

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
            regExLabel: "RegEx");

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
            regExLabel: "RegEx");

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

            var methodName = $"Display{commandParameter}VB";

            // NOTE(crhodes)
            // This expects the method to be public.  Research how to find private methods
            MethodInfo displayWalkerMethod = this.GetType().GetMethod(methodName);
            SearchTreeCommand walkerMethodDelegate = (SearchTreeCommand)displayWalkerMethod.CreateDelegate(typeof(SearchTreeCommand), this);
            EventAggregator.GetEvent<InvokeVBSyntaxWalkerEvent>().Publish(walkerMethodDelegate);

            Log.EVENT("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public bool WalkerCanExecute(string tag)
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #region THESE ARE GOOD

        public StringBuilder DisplayImportsStatementWalkerVB(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.ImportsStatement();

            commandConfiguration.WalkerPattern = ImportsStatementWalker;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayNamespaceStatementWalkerVB(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.NamespaceStatement();

            commandConfiguration.WalkerPattern = NamespaceStatementWalker;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayClassStatementWalkerVB(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            VNCSW.VB.VNCVBSyntaxWalkerBase walker = null;

            commandConfiguration.WalkerPattern = ClassStatementWalker;

            if (ClassStatementWalker.DisplayBlock)
            {
                walker = new VNCSW.VB.ClassBlock();
            }
            else
            {
                walker = new VNCSW.VB.ClassStatement();
            }

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayModuleStatementWalkerVB(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            VNCSW.VB.VNCVBTypedSyntaxWalkerBase walker = null;

            commandConfiguration.WalkerPattern = ModuleStatementWalker;

            if (ModuleStatementWalker.DisplayBlock)
            {
                walker = new VNCSW.VB.ModuleBlock();
            }
            else
            {
                walker = new VNCSW.VB.ModuleStatement();
            }

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayPropertyStatementWalkerVB(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            VNCSW.VB.VNCVBTypedSyntaxWalkerBase walker = null;

            commandConfiguration.WalkerPattern = PropertyStatementWalker;

            if (PropertyStatementWalker.DisplayBlock)
            {
                walker = new VNCSW.VB.PropertyBlock();
            }
            else
            {
                walker = new VNCSW.VB.PropertyStatement();
            }

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayMethodStatementWalkerVB(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            VNCSW.VB.VNCVBTypedSyntaxWalkerBase walker = null;

            commandConfiguration.WalkerPattern = MethodStatementWalker;

            if (MethodStatementWalker.DisplayBlock)
            {
                walker = new VNCSW.VB.MethodBlock();
            }
            else
            {
                walker = new VNCSW.VB.MethodStatement();
            }

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayHandlesClauseWalkerVB(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.HandlesClause();

            commandConfiguration.WalkerPattern = HandlesClauseWalker;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayParameterListWalkerVB(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.ParameterList();

            commandConfiguration.WalkerPattern = ParameterListWalker;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayLocalDeclarationStatementWalkerVB(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.LocalDeclarationStatement();

            commandConfiguration.WalkerPattern = LocalDeclarationStatementWalker;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        #endregion

        #region THIS IS BEING TESTED

        public StringBuilder DisplayMultiLineLambdaExpressionVB(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.MultiLineLambdaExpression();

            commandConfiguration.WalkerPattern = MultiLineLambdaExpressionWalker;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplaySingleLineLambdaExpressionVB(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.SingleLineLambdaExpression();

            commandConfiguration.WalkerPattern = SingleLineLambdaExpressionWalker;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayExpressionStatementVB(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.ExpressionStatement();

            commandConfiguration.WalkerPattern = ExpressionStatementWalker;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayInvocationExpressionWalkerVB(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.InvocationExpression();

            commandConfiguration.WalkerPattern = InvocationExpressionWalker;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayMemberAccessExpressionWalkerVB(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.MemberAccessExpression();

            commandConfiguration.WalkerPattern = MemberAccessExpressionWalker;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayArgumentListWalkerVB(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.ArgumentList();

            commandConfiguration.WalkerPattern = ArgumentListWalker;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayVariableDeclaratorWalkerVB(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.VariableDeclarator();

            commandConfiguration.WalkerPattern = VariableDeclaratorWalker;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayObjectCreationExpressionWalkerVB(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.ObjectCreationExpression();

            commandConfiguration.WalkerPattern = ObjectCreationExpressionWalker;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayBinaryExpressiontWalkerVB(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.BinaryExpression();

            commandConfiguration.WalkerPattern = BinaryExpressionWalker;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayAsNewClauseVB(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.AsNewClause();

            commandConfiguration.WalkerPattern = AsNewClauseWalker;

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplaySimpleAsClauseWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {

            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);
            var walker = new VNCSW.VB.SimpleAsClause();

            commandConfiguration.WalkerPattern = SimpleAsClauseWalker;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplaySyntaxNodeWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.SyntaxNode();

            commandConfiguration.WalkerPattern = SyntaxNodeWalker;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplaySyntaxTokenWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.SyntaxToken();

            commandConfiguration.WalkerPattern = SyntaxTokenWalker;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplaySyntaxTriviaWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.SyntaxTrivia();

            commandConfiguration.WalkerPattern = SyntaxTriviaWalker;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        #endregion

        #region THESE NEED TO BE UPDATED


        public StringBuilder DisplayStopOrEndStatementVB(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.StopOrEndStatement();

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayAssignmentStatementWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.AssignmentStatement();

            //walker.MatchLeft = (bool)ceAssignmentStatementMatchLeft.IsChecked;
            //walker.MatchRight = (bool)ceAssignmentStatementMatchRight.IsChecked;

            //commandConfiguration.UseRegEx = (bool)ceAssignmentStatementUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teAssignmentStatementRegEx.Text;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        private StringBuilder DisplayStructureBlockWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.StructureBlock();

            //walker.ShowFields = (bool)ceStructureShowFields.IsChecked;

            //walker.HasAttributes = (bool)CodeExplorer.configurationOptions.ceHasAttributes.IsChecked;

            //walker.AllFieldTypes = (bool)CodeExplorer.configurationOptions.ceAllTypes.IsChecked;
            //walker.FieldNames = (bool)ceStructureFieldsUseRegEx.IsChecked ? teStructureFieldsRegEx.Text : ".*";
            //walker.StructureNames = (bool)ceStructuresUseRegEx.IsChecked ? teStructureRegEx.Text : ".*";

            //commandConfiguration.UseRegEx = (bool)ceStructuresUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teStructureRegEx.Text;

            // StructureBlock has special (two types) of RegEx.
            walker.InitializeRegEx();

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayFieldDeclarationWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            VNCCA.SyntaxNode.FieldDeclarationLocation fieldDeclarationLocation = VNCCA.SyntaxNode.FieldDeclarationLocation.Class;

            // TODO(crhodes)
            // Go look at EyeOnLife and see how to do this in a cleaner way.

            //switch (lbeFieldDeclarationLocation.EditValue.ToString())
            //{
            //    case "Class":
            //        fieldDeclarationLocation = VNCCA.SyntaxNode.FieldDeclarationLocation.Class;
            //        break;

            //    case "Module":
            //        fieldDeclarationLocation = VNCCA.SyntaxNode.FieldDeclarationLocation.Module;
            //        break;

            //    case "Structure":
            //        fieldDeclarationLocation = VNCCA.SyntaxNode.FieldDeclarationLocation.Structure;
            //        break;
            //}

            VNCSW.VB.VNCVBTypedSyntaxWalkerBase walker = null;

            walker = new VNCSW.VB.FieldDeclaration(fieldDeclarationLocation);

            //walker.HasAttributes = (bool)CodeExplorer.configurationOptions.ceHasAttributes.IsChecked;

            //commandConfiguration.UseRegEx = (bool)ceFieldDeclarationUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teFieldDeclarationRegEx.Text;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
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
