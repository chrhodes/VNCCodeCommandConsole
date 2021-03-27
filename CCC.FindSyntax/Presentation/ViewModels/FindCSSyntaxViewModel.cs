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

            NamespaceDeclarationWalkerCommand = new DelegateCommand(
                NamespaceDeclarationWalkerExecute, NamespaceDeclarationWalkerCanExecute);

            ClassDeclarationWalkerCommand = new DelegateCommand(
                ClassDeclarationWalkerExecute, ClassDeclarationWalkerCanExecute);

            FieldDeclarationWalkerCommand = new DelegateCommand(
                FieldDeclarationWalkerExecute, FieldDeclarationWalkerCanExecute);







        Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Enums


        #endregion

        #region Structures


        #endregion

        #region Fields and Properties


        private string _declarationLocation;
        private string _usingDirectiveRegEx = ".*";
        private bool _usingDirectiveUseRegEx;

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

        private string _namespaceDeclarationRegEx = ".*";
        private bool _namespaceDeclarationUseRegEx;

        public string NamespaceDeclarationRegEx
        {
            get => _namespaceDeclarationRegEx;
            set
            {
                if (_namespaceDeclarationRegEx == value)
                    return;
                _namespaceDeclarationRegEx = value;
                OnPropertyChanged();
            }
        }

        public bool NamespaceDeclarationUseRegEx
        {
            get => _namespaceDeclarationUseRegEx;
            set
            {
                if (_namespaceDeclarationUseRegEx == value)
                    return;
                _namespaceDeclarationUseRegEx = value;
                OnPropertyChanged();
            }
        }

        private string _classDeclarationRegEx = ".*";
        private bool _classDeclarationUseRegEx;

        public string ClassDeclarationRegEx
        {
            get => _classDeclarationRegEx;
            set
            {
                if (_classDeclarationRegEx == value)
                    return;
                _classDeclarationRegEx = value;
                OnPropertyChanged();
            }
        }

        public bool ClassDeclarationUseRegEx
        {
            get => _classDeclarationUseRegEx;
            set
            {
                if (_classDeclarationUseRegEx == value)
                    return;
                _classDeclarationUseRegEx = value;
                OnPropertyChanged();
            }
        }

        private string _fieldDeclarationRegEx;
        private bool _fieldDeclarationUseRegEx;

        public string FieldDeclarationRegEx
        {
            get => _fieldDeclarationRegEx;
            set
            {
                if (_fieldDeclarationRegEx == value)
                    return;
                _fieldDeclarationRegEx = value;
                OnPropertyChanged();
            }
        }

        public bool FieldDeclarationUseRegEx
        {
            get => _fieldDeclarationUseRegEx;
            set
            {
                if (_fieldDeclarationUseRegEx == value)
                    return;
                _fieldDeclarationUseRegEx = value;
                OnPropertyChanged();
            }
        }

        
        public string DeclarationLocation
        {
            get => _declarationLocation;
            set
            {
                if (_declarationLocation == value)
                    return;
                _declarationLocation = value;
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

        #region ClassDeclaration Walker

        public DelegateCommand ClassDeclarationWalkerCommand { get; set; }

        public void ClassDeclarationWalkerExecute()
        {
            Int64 startTicks = Log.EVENT("Enter", Common.LOG_CATEGORY);

            EventAggregator.GetEvent<InvokeCSSyntaxWalkerEvent>().Publish(DisplayClassDeclarationWalkerCS);

            Log.EVENT("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public bool ClassDeclarationWalkerCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        private StringBuilder DisplayClassDeclarationWalkerCS(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.CS.ClassDeclaration();

            commandConfiguration.UseRegEx = ClassDeclarationUseRegEx;
            commandConfiguration.RegEx = ClassDeclarationRegEx;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.CS.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        #endregion

        // Begin Cut 3
        // Put this in Commands, typically Private or Public Methods

        #region FieldDeclaration Walker

        public DelegateCommand FieldDeclarationWalkerCommand { get; set; }

        public void FieldDeclarationWalkerExecute()
        {
            Int64 startTicks = Log.EVENT("Enter", Common.LOG_CATEGORY);

            EventAggregator.GetEvent<InvokeCSSyntaxWalkerEvent>().Publish(DisplayFieldDeclarationWalkerCS);

            Log.EVENT("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public bool FieldDeclarationWalkerCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        private StringBuilder DisplayFieldDeclarationWalkerCS(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            VNCCA.SyntaxNode.FieldDeclarationLocation fieldDeclarationLocation = VNCCA.SyntaxNode.FieldDeclarationLocation.Class;

            // TODO(crhodes)
            // Look at VB Code

            var walker = new VNCSW.CS.FieldDeclaration(fieldDeclarationLocation);

            commandConfiguration.UseRegEx = FieldDeclarationUseRegEx;
            commandConfiguration.RegEx = FieldDeclarationRegEx;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.CS.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        #endregion

        // End Cut 3
        #region NamespaceDeclaration Walker

        public DelegateCommand NamespaceDeclarationWalkerCommand { get; set; }

        public void NamespaceDeclarationWalkerExecute()
        {
            Int64 startTicks = Log.EVENT("Enter", Common.LOG_CATEGORY);

            EventAggregator.GetEvent<InvokeCSSyntaxWalkerEvent>().Publish(DisplayNamespaceDeclarationWalkerCS);

            Log.EVENT("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public bool NamespaceDeclarationWalkerCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        private StringBuilder DisplayNamespaceDeclarationWalkerCS(SearchTreeCommandConfiguration commandConfiguration)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var walker = new VNCSW.CS.NamespaceDeclaration();

            commandConfiguration.UseRegEx = NamespaceDeclarationUseRegEx;
            commandConfiguration.RegEx = NamespaceDeclarationRegEx;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return VNCCA.Helpers.CS.InvokeVNCSyntaxWalker(walker, commandConfiguration);
        }

        #endregion

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
