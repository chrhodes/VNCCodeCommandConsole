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

            ImportsStatementWalkerCommand = new DelegateCommand(
                ImportsStatementWalkerExecute, ImportsStatementWalkerCanExecute);

            NamespaceStatementWalkerCommand = new DelegateCommand(
                NamespaceStatementWalkerExecute, NamespaceStatementWalkerCanExecute);

            ModuleStatementWalkerCommand = new DelegateCommand(
                ModuleStatementWalkerExecute, ModuleStatementWalkerCanExecute);

            SyntaxWalkerCommand = new DelegateCommand(
    WalkerExecute, WalkerCanExecute);

            // TODO(crhodes)
            // Either add the methods or write something reflective on the tag
            // to call the correct method.
            //
            // Longer term see if can make the ones that just support RegEx and UseRegEx simpler.

            //FieldDeclarationWalkerCommand = new DelegateCommand(
            //    FieldDeclarationWalkerExecute, FieldDeclarationWalkerCanExecute);

            //PropertyStatementWalkerCommand = new DelegateCommand(
            //    PropertyStatementWalkerExecute, PropertyStatementWalkerCanExecute);

            //FindStructureBlockWalkerCommand = new DelegateCommand(
            //    FindStructureBlockWalkerExecute, FindStructureBlockWalkerCanExecute);

            //MethodBlockWalkerCommand = new DelegateCommand(
            //    MethodBlockWalkerExecute, MethodBlockWalkerCanExecute);

            //HandlesClauseWalkerCommand = new DelegateCommand(
            //    HandlesClauseWalkerExecute, HandlesClauseWalkerCanExecute);

            //ParameterListWalkerCommand = new DelegateCommand(
            //    ParameterListWalkerExecute, ParameterListWalkerCanExecute);

            //LocalDelcarationStatementWalkerCommand = new DelegateCommand(
            //    LocalDelcarationStatementWalkerExecute, LocalDelcarationStatementWalkerCanExecute);

            //MultiLineLambdaExpressionWalkerCommand = new DelegateCommand(
            //    MultiLineLambdaExpressionWalkerExecute, MultiLineLambdaExpressionWalkerCanExecute);

            //SingleLineLambdaExpressionWalkerCommand = new DelegateCommand(
            //    SingleLineLambdaExpressionWalkerCommandExecute, SingleLineLambdaExpressionWalkerCommandCanExecute);

            //ExpressionStatementWalkerCommand = new DelegateCommand(
            //    ExpressionStatementWalkerExecute, ExpressionStatementWalkerCanExecute);

            //InvocationExpressionWalkerCommand = new DelegateCommand(
            //    InvocationExpressionWalkerExecute, InvocationExpressionWalkerCanExecute);

            //MemberAccessExpressionWalkerCommand = new DelegateCommand(
            //    MemberAccessExpressionWalkerExecute, MemberAccessExpressionWalkerCanExecute);

            //ArgumentListWalkerCommand = new DelegateCommand(
            //    ArgumentListWalkerExecute, ArgumentListWalkerCanExecute);

            //VariableDeclaratorWalkerCommand = new DelegateCommand(
            //    VariableDeclaratorWalkerExecute, VariableDeclaratorWalkerCanExecute);

            //ObjectCreationExpressionWalkerCommand = new DelegateCommand(
            //    ObjectCreationExpressionWalkerExecute, ObjectCreationExpressionWalkerCanExecute);

            //AssignmentStatementWalkerCommand = new DelegateCommand(
            //    AssignmentStatementWalkerExecute, AssignmentStatementWalkerCanExecute);

            //BinaryExpressionWalkerCommand = new DelegateCommand(
            //    BinaryExpressionWalkerExecute, BinaryExpressionWalkerCanExecute);

            //AsNewClauseWalkerCommand = new DelegateCommand(
            //    AsNewClauseWalkerExecute, AsNewClauseWalkerCanExecute);

            //SimpleAsClauseWalkerCommand = new DelegateCommand(
            //    SimpleAsClauseWalkerExecute, SimpleAsClauseWalkerCanExecute);

            //StopOrEndStatementWalkerCommand = new DelegateCommand(
            //    StopOrEndStatementWalkerExecute, StopOrEndStatementWalkerCanExecute);

            //SyntaxNodeWalkerCommand = new DelegateCommand(
            //    SyntaxNodeWalkerExecute, SyntaxNodeWalkerCanExecute);

            //SyntaxTokenWalkerCommand = new DelegateCommand(
            //    SyntaxTokenWalkerExecute, SyntaxTokenWalkerCanExecute);

            //SyntaxTriviaWalkerCommand = new DelegateCommand(
            //    SyntaxTriviaWalkerExecute, SyntaxTriviaWalkerCanExecute);

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

    #endregion

    #region Enums


    #endregion

    #region Structures


    #endregion

    #region Fields and Properties

    private string _namespaceStatementRegEx;
    private bool _namespaceStatementUseRegEx;

    public string NamespaceStatementRegEx
    {
        get => _namespaceStatementRegEx;
        set
        {
            if (_namespaceStatementRegEx == value)
                return;
            _namespaceStatementRegEx = value;
            OnPropertyChanged();
        }
    }

    public bool NamespaceStatementUseRegEx
    {
        get => _namespaceStatementUseRegEx;
        set
        {
            if (_namespaceStatementUseRegEx == value)
                return;
            _namespaceStatementUseRegEx = value;
            OnPropertyChanged();
        }
    }

    private bool _importsStatementUseRegEx;
        private string _importsStatementRegEx = ".*";

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

        private string _importsStatementRegEx2 = "99";

        public string ImportsStatementRegEx2
        {
            get => _importsStatementRegEx2;
            set
            {
                if (_importsStatementRegEx2 == value)
                    return;
                _importsStatementRegEx2 = value;
                OnPropertyChanged();
            }
        }

        private string _moduleStatementRegEx;
    private bool _moduleStatementUseRegEx;

    public string ModuleStatementRegEx
    {
        get => _moduleStatementRegEx;
        set
        {
            if (_moduleStatementRegEx == value)
                return;
            _moduleStatementRegEx = value;
            OnPropertyChanged();
        }
    }

    public bool ModuleStatementUseRegEx
    {
        get => _moduleStatementUseRegEx;
        set
        {
            if (_moduleStatementUseRegEx == value)
                return;
            _moduleStatementUseRegEx = value;
            OnPropertyChanged();
        }
    }

    private bool _expressionStatementUseRegEx;
        private string _expressionStatementRegEx = ".*";

        public string ExpressionStatementRegEx
        {
            get => _expressionStatementRegEx;
            set
            {
                if (_expressionStatementRegEx == value)
                    return;
                _expressionStatementRegEx = value;
                OnPropertyChanged();
            }
        }

        public bool ExpressionStatementUseRegEx
        {
            get => _expressionStatementUseRegEx;
            set
            {
                if (_expressionStatementUseRegEx == value)
                    return;
                _expressionStatementUseRegEx = value;
                OnPropertyChanged();
            }
        }

        private bool _handlesClauseUseRegEx;
        private string _handlesClauseRegEx = ".*";

        public string HandlesClauseRegEx
        {
            get => _handlesClauseRegEx;
            set
            {
                if (_handlesClauseRegEx == value)
                    return;
                _handlesClauseRegEx = value;
                OnPropertyChanged();
            }
        }

        public bool HandlesClauseUseRegEx
        {
            get => _handlesClauseUseRegEx;
            set
            {
                if (_handlesClauseUseRegEx == value)
                    return;
                _handlesClauseUseRegEx = value;
                OnPropertyChanged();
            }
        }

        private string _message = "Initial Message";
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

        #region Event Handlers

        #endregion

        #region Public Methods


        #endregion

        #region Protected Methods


        #endregion

        #region Private Methods

        public DelegateCommand SyntaxWalkerCommand { get; set; }

        public void WalkerExecute()
        {
            Int64 startTicks = Log.EVENT("Enter", Common.LOG_CATEGORY);

            //Helper.ProcessOperation(DisplayImportsStatementWalkerVB, CodeExplorer, CodeExplorerContext, CodeExplorer.configurationOptions);

            Message = $"Time is {DateTime.Now}";
            var foo = ImportsStatementRegEx2;

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

            //Helper.ProcessOperation(DisplayImportsStatementWalkerVB, CodeExplorer, CodeExplorerContext, CodeExplorer.configurationOptions);

            var foo = ImportsStatementRegEx2;
            EventAggregator.GetEvent<InvokeVBSyntaxWalkerEvent>().Publish(DisplayImportsStatementWalkerVB);

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

        #region NamespaceStatement Walker

        public DelegateCommand NamespaceStatementWalkerCommand { get; set; }

        public void NamespaceStatementWalkerExecute()
        {
            Int64 startTicks = Log.EVENT("Enter", Common.LOG_CATEGORY);

            EventAggregator.GetEvent<InvokeVBSyntaxWalkerEvent>().Publish(DisplayNamespaceStatementWalkerVB);

            Log.EVENT("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public bool NamespaceStatementWalkerCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        private StringBuilder DisplayNamespaceStatementWalkerCS(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.NamespaceStatement();

            commandConfiguration.UseRegEx = NamespaceStatementUseRegEx;
            commandConfiguration.RegEx = NamespaceStatementRegEx;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        #endregion

        #region ModuleStatement Walker

        public DelegateCommand ModuleStatementWalkerCommand { get; set; }

        public void ModuleStatementWalkerExecute()
        {
            Int64 startTicks = Log.EVENT("Enter", Common.LOG_CATEGORY);

            EventAggregator.GetEvent<InvokeVBSyntaxWalkerEvent>().Publish(DisplayModuleStatementWalkerVB);

            Log.EVENT("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public bool ModuleStatementWalkerCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        private StringBuilder DisplayModuleStatementWalkerCS(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.ModuleStatement();

            commandConfiguration.UseRegEx = ModuleStatementUseRegEx;
            commandConfiguration.RegEx = ModuleStatementRegEx;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        #endregion

        #region THIS IS GOOD

        private StringBuilder DisplayImportsStatementWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.ImportsStatement();

            commandConfiguration.UseRegEx = ImportsStatementUseRegEx;
            commandConfiguration.RegEx = ImportsStatementRegEx;

            commandConfiguration.RegEx = ImportsStatementRegEx;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        #endregion

        StringBuilder DisplayStopOrEndStatementVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.StopOrEndStatement();

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        StringBuilder DisplayExpressionStatementVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.ExpressionStatement();

            commandConfiguration.UseRegEx = ExpressionStatementUseRegEx;
            commandConfiguration.RegEx = ExpressionStatementRegEx;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        StringBuilder DisplayHandlesClauseVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {

            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);
            var walker = new VNCSW.VB.HandlesClause();

            commandConfiguration.UseRegEx = HandlesClauseUseRegEx;
            commandConfiguration.RegEx = HandlesClauseRegEx;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        StringBuilder DisplayMultiLineLambdaExpressionVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.MultiLineLambdaExpression();

            //commandConfiguration.UseRegEx = (bool)ceMultiLineLambdaExpressionUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teMultiLineLambdaExpressionRegEx.Text;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        StringBuilder DisplaySingleLineLambdaExpressionVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.SingleLineLambdaExpression();

            //commandConfiguration.UseRegEx = (bool)ceSingleLineLambdaExpressionUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teSingleLineLambdaExpressionRegEx.Text;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        private StringBuilder DisplayMemberAccessExpressionWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.MemberAccessExpression();

            //commandConfiguration.UseRegEx = (bool)ceMemberAccessExpressionUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teMemberAccessExpressionRegEx.Text;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        StringBuilder DisplaySyntaxNodeWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.SyntaxNode();

            //commandConfiguration.UseRegEx = (bool)ceSyntaxNodeUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teSyntaxNodeRegEx.Text;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        StringBuilder DisplaySyntaxTokenWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.SyntaxToken();

            //commandConfiguration.UseRegEx = (bool)ceSyntaxTokenUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teSyntaxTokenRegEx.Text;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        StringBuilder DisplayAsNewClauseVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.AsNewClause();

            //commandConfiguration.UseRegEx = (bool)ceAsNewClauseUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teAsNewClauseRegEx.Text;

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        StringBuilder DisplaySimpleAsClauseWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {

            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);
            var walker = new VNCSW.VB.SimpleAsClause();

            //commandConfiguration.UseRegEx = (bool)ceSimpleAsClauseUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teSimpleAsClauseRegEx.Text;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        StringBuilder DisplayObjectCreationExpressionWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.ObjectCreationExpression();

            //commandConfiguration.UseRegEx = (bool)ceObjectCreationExpressionUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teObjectCreationExpressionRegEx.Text;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        StringBuilder DisplaySyntaxTriviaWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.SyntaxTrivia();

            //commandConfiguration.UseRegEx = (bool)ceSyntaxTriviaUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teSyntaxTriviaRegEx.Text;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        private StringBuilder DisplayBinaryExpressiontWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.BinaryExpression();

            //commandConfiguration.UseRegEx = (bool)ceBinaryExpressionUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teBinaryExpressionRegEx.Text;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }
        StringBuilder DisplayAssignmentStatementWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
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

        StringBuilder DisplayLocalDeclarationStatementWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.LocalDeclarationStatement();

            //walker.HasAttributes = (bool)CodeExplorer.configurationOptions.ceHasAttributes.IsChecked;

            //commandConfiguration.UseRegEx = (bool)ceLocalDeclarationStatementUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teLocalDeclarationStatementRegEx.Text;

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

        private StringBuilder DisplayVariableDeclaratorWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.VariableDeclarator();

            //walker.HasAttributes = (bool)CodeExplorer.configurationOptions.ceHasAttributes.IsChecked;

            //commandConfiguration.UseRegEx = (bool)ceVariableDeclaratorUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teVariableDeclaratorRegEx.Text;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        private StringBuilder DisplayInvocationExpressionWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.InvocationExpression();

            //commandConfiguration.UseRegEx = (bool)ceInvocationExpressionUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teInvocationExpressionRegEx.Text;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        private StringBuilder DisplayParameterListWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.ParameterList();

            //commandConfiguration.UseRegEx = (bool)ceParameterListUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teParameterListRegEx.Text;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        private StringBuilder DisplayArgumentListWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.ArgumentList();

            //commandConfiguration.UseRegEx = (bool)ceArgumentListUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teArgumentListRegEx.Text;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        private StringBuilder DisplayPropertyStatementWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            VNCSW.VB.VNCVBTypedSyntaxWalkerBase walker = null;

            //if ((bool)ceShowPropertyBlock.IsChecked)
            //{
            //    walker = new VNCSW.VB.PropertyBlock();
            //}
            //else
            //{
            //    walker = new VNCSW.VB.PropertyStatement();
            //}

            //walker.HasAttributes = (bool)CodeExplorer.configurationOptions.ceHasAttributes.IsChecked;

            //commandConfiguration.UseRegEx = (bool)cePropertyStatementUseRegEx.IsChecked;
            //commandConfiguration.RegEx = tePropertyStatementRegEx.Text;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        private StringBuilder DisplayFieldDeclarationWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
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

        private StringBuilder DisplayClassStatementWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            VNCSW.VB.VNCVBSyntaxWalkerBase walker = null;

            //if ((bool)ceShowClassBlock.IsChecked)
            //{
            //    walker = new VNCSW.VB.ClassBlock();
            //}
            //else
            //{
            //    walker = new VNCSW.VB.ClassStatement();
            //}

            //commandConfiguration.UseRegEx = (bool)ceClassStatementUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teClassStatementRegEx.Text;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        StringBuilder DisplayMethodBlockWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            VNCSW.VB.VNCVBTypedSyntaxWalkerBase walker = null;

            //commandConfiguration.UseRegEx = (bool)ceMethodBlockUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teMethodBlockRegEx.Text;

            // TODO(crhodes)
            // Maybe figure out how to suppress showing of block.

            //if ((bool)ceShowMethodBlock2.IsChecked)
            //{
            walker = new VNCSW.VB.MethodBlock();
            //commandConfiguration.ConfigurationOptions.ShowAnalysisCRC = true;
            //}
            //else
            //{
            //    walker = new VNCSW.VB.MethodStatement();
            //}

            StringBuilder results = VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);

            // We may have done a deep dive on a method.  Go grab the results.
            // TODO(crhodes)
            // This might only be if in MethodBlock mode.  See above.

            //CodeExplorer.teSyntaxNode.Text += walker.WalkerNode.ToString();
            //CodeExplorer.teSyntaxToken.Text += walker.WalkerToken.ToString();
            //CodeExplorer.teSyntaxTrivia.Text += walker.WalkerTrivia.ToString();
            //CodeExplorer.teSyntaxStructuredTrivia.Text += walker.WalkerStructuredTrivia.ToString();

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return results;
        }

        private StringBuilder DisplayMethodStatementWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            VNCSW.VB.VNCVBTypedSyntaxWalkerBase walker = null;

            //commandConfiguration.UseRegEx = (bool)ceMethodStatementUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teMethodStatementRegEx.Text;

            //if ((bool)ceShowMethodBlock.IsChecked)
            //{
            //    walker = new VNCSW.VB.MethodBlock();
            //    //commandConfiguration.ConfigurationOptions.ShowAnalysisCRC = true;
            //}
            //else
            //{
            //    walker = new VNCSW.VB.MethodStatement();
            //}

            StringBuilder results = VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);

            // We may have done a deep dive on a method.  Go grab the results.
            // TODO(crhodes)
            // This might only be if in MethodBlock mode.  See above.

            //CodeExplorer.teSyntaxNode.Text += walker.WalkerNode.ToString();
            //CodeExplorer.teSyntaxToken.Text += walker.WalkerToken.ToString();
            //CodeExplorer.teSyntaxTrivia.Text += walker.WalkerTrivia.ToString();
            //CodeExplorer.teSyntaxStructuredTrivia.Text += walker.WalkerStructuredTrivia.ToString();

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return results;
        }

        private StringBuilder DisplayModuleStatementWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            VNCSW.VB.VNCVBTypedSyntaxWalkerBase walker = null;

            //if ((bool)ceShowModuleBlock.IsChecked)
            //{
            //    walker = new VNCSW.VB.ModuleBlock();
            //}
            //else
            //{
            //    walker = new VNCSW.VB.ModuleStatement();
            //}

            //commandConfiguration.UseRegEx = (bool)ceModuleStatementUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teModuleStatementRegEx.Text;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        private StringBuilder DisplayNamespaceStatementWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.NamespaceStatement();

            //commandConfiguration.UseRegEx = (bool)ceNamespaceStatementUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teNamespaceStatementRegEx.Text;

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
