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

        #region Enums


        #endregion

        #region Structures


        #endregion

        #region Fields and Properties

        // This does not work.  Needs to be a property.
        //WalkerPattern ImportsStatementWalker = new WalkerPattern();

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
            // from the WalkerPattern property that correspondes to the walkerPropertyName

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


        #endregion

        #region THIS IS BEING TESTED

        public StringBuilder DisplayMethodStatementWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
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

        public StringBuilder DisplayMethodBlockWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
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




        #endregion

        #region THESE NEED TO BE UPDATED


        public StringBuilder DisplayStopOrEndStatementVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.StopOrEndStatement();

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayExpressionStatementVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.ExpressionStatement();

            //commandConfiguration.WalkerPattern.UseRegEx = ExpressionStatementUseRegEx;
            //commandConfiguration.WalkerPattern.RegEx = ExpressionStatementRegEx;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayHandlesClauseVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {

            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);
            var walker = new VNCSW.VB.HandlesClause();

            //commandConfiguration.WalkerPattern.UseRegEx = HandlesClauseUseRegEx;
            //commandConfiguration.WalkerPattern.RegEx = HandlesClauseRegEx;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayMultiLineLambdaExpressionVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.MultiLineLambdaExpression();

            //commandConfiguration.UseRegEx = (bool)ceMultiLineLambdaExpressionUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teMultiLineLambdaExpressionRegEx.Text;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplaySingleLineLambdaExpressionVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.SingleLineLambdaExpression();

            //commandConfiguration.UseRegEx = (bool)ceSingleLineLambdaExpressionUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teSingleLineLambdaExpressionRegEx.Text;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayMemberAccessExpressionWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.MemberAccessExpression();

            //commandConfiguration.UseRegEx = (bool)ceMemberAccessExpressionUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teMemberAccessExpressionRegEx.Text;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplaySyntaxNodeWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.SyntaxNode();

            //commandConfiguration.UseRegEx = (bool)ceSyntaxNodeUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teSyntaxNodeRegEx.Text;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplaySyntaxTokenWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.SyntaxToken();

            //commandConfiguration.UseRegEx = (bool)ceSyntaxTokenUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teSyntaxTokenRegEx.Text;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayAsNewClauseVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.AsNewClause();

            //commandConfiguration.UseRegEx = (bool)ceAsNewClauseUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teAsNewClauseRegEx.Text;

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplaySimpleAsClauseWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {

            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);
            var walker = new VNCSW.VB.SimpleAsClause();

            //commandConfiguration.UseRegEx = (bool)ceSimpleAsClauseUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teSimpleAsClauseRegEx.Text;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayObjectCreationExpressionWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.ObjectCreationExpression();

            //commandConfiguration.UseRegEx = (bool)ceObjectCreationExpressionUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teObjectCreationExpressionRegEx.Text;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplaySyntaxTriviaWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.SyntaxTrivia();

            //commandConfiguration.UseRegEx = (bool)ceSyntaxTriviaUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teSyntaxTriviaRegEx.Text;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayBinaryExpressiontWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.BinaryExpression();

            //commandConfiguration.UseRegEx = (bool)ceBinaryExpressionUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teBinaryExpressionRegEx.Text;

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

        public StringBuilder DisplayLocalDeclarationStatementWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
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

        public StringBuilder DisplayVariableDeclaratorWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.VariableDeclarator();

            //walker.HasAttributes = (bool)CodeExplorer.configurationOptions.ceHasAttributes.IsChecked;

            //commandConfiguration.UseRegEx = (bool)ceVariableDeclaratorUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teVariableDeclaratorRegEx.Text;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayInvocationExpressionWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.InvocationExpression();

            //commandConfiguration.UseRegEx = (bool)ceInvocationExpressionUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teInvocationExpressionRegEx.Text;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayParameterListWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.ParameterList();

            //commandConfiguration.UseRegEx = (bool)ceParameterListUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teParameterListRegEx.Text;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.VB.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        public StringBuilder DisplayArgumentListWalkerVB(VNCCA.SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.VB.ArgumentList();

            //commandConfiguration.UseRegEx = (bool)ceArgumentListUseRegEx.IsChecked;
            //commandConfiguration.RegEx = teArgumentListRegEx.Text;

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
