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

using VNCCA = VNC.CodeAnalysis;
using VNCSW = VNC.CodeAnalysis.SyntaxWalkers;

namespace VNCCodeCommandConsole.Presentation.ViewModels
{
    public class ConfigurationOptionsViewModel : EventViewModelBase, IConfigurationOptionsViewModel, IInstanceCountVM
    {

        #region Constructors, Initialization, and Load

        public ConfigurationOptionsViewModel(
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

            Message = "ConfigurationOptionsViewModel says hello";

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Enums


        #endregion

        #region Structures


        #endregion

        #region Fields and Properties

        public ICommand SayHelloCommand { get; private set; }

        private bool _displayCRC32;
        private int _displaySummaryMinimum;
        private bool _displaySummary = true;
        private bool _displayResults = true;
        private bool _alwaysDisplayFileName;
        private bool _listImpactedFilesOnly;
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


        public bool ListImpactedFilesOnly
        {
            get => _listImpactedFilesOnly;
            set
            {
                if (_listImpactedFilesOnly == value)
                    return;
                _listImpactedFilesOnly = value;
                OnPropertyChanged();
            }
        }


        public bool AlwaysDisplayFileName
        {
            get => _alwaysDisplayFileName;
            set
            {
                if (_alwaysDisplayFileName == value)
                    return;
                _alwaysDisplayFileName = value;
                OnPropertyChanged();
            }
        }


        public bool DisplayResults
        {
            get => _displayResults;
            set
            {
                if (_displayResults == value)
                    return;
                _displayResults = value;
                OnPropertyChanged();
            }
        }


        public bool DisplaySummary
        {
            get => _displaySummary;
            set
            {
                if (_displaySummary == value)
                    return;
                _displaySummary = value;
                OnPropertyChanged();
            }
        }


        public int DisplaySummaryMinimum
        {
            get => _displaySummaryMinimum;
            set
            {
                if (_displaySummaryMinimum == value)
                    return;
                _displaySummaryMinimum = value;
                OnPropertyChanged();
            }
        }

        
        public bool DisplayCRC32
        {
            get => _displayCRC32;
            set
            {
                if (_displayCRC32 == value)
                    return;
                _displayCRC32 = value;
                OnPropertyChanged();
            }
        }
        




        #endregion

        #region Event Handlers



        #endregion

        #region Public Methods

        public VNC.CodeAnalysis.SyntaxConfigurationOptions GetSyntaxConfigurationOptions()
        {
            VNC.CodeAnalysis.SyntaxConfigurationOptions configurationOptions = new VNCCA.SyntaxConfigurationOptions();

            // TODO(crhodes)
            // Bind all this to new Properties in VM

            //var foo = lbeSyntaxWalkerDepth.EditValue;
            //var bar = lbeAdditionalNodes.EditValue;

            //switch (lbeSyntaxWalkerDepth.EditValue)
            //{
            //    case "Node":
            //        configurationOptions.SyntaxWalkerDepth = Microsoft.CodeAnalysis.SyntaxWalkerDepth.Node;
            //        break;

            //    case "Token":
            //        configurationOptions.SyntaxWalkerDepth = Microsoft.CodeAnalysis.SyntaxWalkerDepth.Token;
            //        break;

            //    case "Trivia":
            //        configurationOptions.SyntaxWalkerDepth = Microsoft.CodeAnalysis.SyntaxWalkerDepth.Trivia;
            //        break;

            //    case "StructureTrivia":
            //        configurationOptions.SyntaxWalkerDepth = Microsoft.CodeAnalysis.SyntaxWalkerDepth.StructuredTrivia;
            //        break;
            //}

            //configurationOptions.AdditionalNodeAnalysis = (VNCCA.SyntaxNode.AdditionalNodes)lbeAdditionalNodes.EditValue;

            //configurationOptions.DisplayNodeKind = (bool)ceDisplay_NodeKind.IsChecked;
            //configurationOptions.DisplayNodeValue = (bool)ceDisplay_NodeValue.IsChecked;
            ////configurationOptions.DisplayFormattedOutput = (bool)ceDisplay_FormattedOutput.IsChecked;

            //configurationOptions.DisplayNodeParent = (bool)ceDisplay_NodeParent.IsChecked;

            //configurationOptions.DisplayStatementBlock = (bool)ceDisplay_StatementBlock.IsChecked;
            //configurationOptions.IncludeStatementBlockInCRC = (bool)ceIncludeStatementBlockInCRC.IsChecked;

            //configurationOptions.SourceLocation = (bool)ceDisplaySourceLocation.IsChecked;
            //configurationOptions.CRC32 = (bool)ceDisplayCRC32.IsChecked;
            //configurationOptions.ReplaceCRLF = (bool)ceReplaceCRLF.IsChecked;

            //configurationOptions.ClassOrModuleName = (bool)ceDisplayClassOrModuleName.IsChecked;
            //configurationOptions.MethodName = (bool)ceDisplayMethodName.IsChecked;

            //configurationOptions.ContainingMethodBlock = (bool)ceDisplayContainingMethodBlock.IsChecked;
            //configurationOptions.ContainingBlock = (bool)ceDisplayContainingBlock.IsChecked;

            //configurationOptions.InTryBlock = (bool)ceInTryBlock.IsChecked;
            //configurationOptions.InWhileBlock = (bool)ceInWhileBlock.IsChecked;
            //configurationOptions.InForBlock = (bool)ceInForBlock.IsChecked;
            //configurationOptions.InIfBlock = (bool)ceInIfBlock.IsChecked;

            //configurationOptions.AllTypes = (bool)ceAllTypes.IsChecked;

            //configurationOptions.Byte = (bool)ceIsByte.IsChecked;
            //configurationOptions.Boolean = (bool)ceIsBoolean.IsChecked;
            //configurationOptions.Date = (bool)ceIsDate.IsChecked;
            //configurationOptions.DataTable = (bool)ceIsDataTable.IsChecked;
            //configurationOptions.DateTime = (bool)ceIsDateTime.IsChecked;
            //configurationOptions.Int16 = (bool)ceIsInt16.IsChecked;
            //configurationOptions.Int32 = (bool)ceIsInt32.IsChecked;
            //configurationOptions.Integer = (bool)ceIsInteger.IsChecked;
            //configurationOptions.Long = (bool)ceIsLong.IsChecked;
            //configurationOptions.Single = (bool)ceIsSingle.IsChecked;
            //configurationOptions.String = (bool)ceIsString.IsChecked;

            //configurationOptions.OtherTypes = (bool)ceOtherTypes.IsChecked;

            //configurationOptions.AddFileSuffix = (bool)ceAddFileSuffix.IsChecked;
            //configurationOptions.FileSuffix = teFileSuffix.Text;

            return configurationOptions;
        }

        #endregion

        #region Protected Methods


        #endregion

        #region Private Methods

        private bool SayHelloCanExecute()
        {
            return true;
        }

        private void SayHello()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Message = "Hello";

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
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
