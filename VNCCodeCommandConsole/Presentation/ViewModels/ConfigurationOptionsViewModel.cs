using System;

using Prism.Events;
using Prism.Services.Dialogs;

using VNC;
using VNC.Core.Mvvm;

using VNCCodeCommandConsole.Presentation.ModelWrappers;

using VNCCA = VNC.CodeAnalysis;

namespace VNCCodeCommandConsole.Presentation.ViewModels
{
    public class ConfigurationOptionsViewModel : EventViewModelBase, IConfigurationOptionsViewModel, IInstanceCountVM
    {
        #region Constructors, Initialization, and Load

        public ConfigurationOptionsViewModel(
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

            CodeAnalysisOptions = new CodeAnalysisOptionsWrapper(
                new VNCCA.CodeAnalysisOptions());

            // NOTE(crhodes)
            // 
            CodeAnalysisOptions.AllTypes = true;
            DisplayContext = false;
            DisplaySummaryMinimum = 1;

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion Constructors, Initialization, and Load

        #region Fields and Properties

        
        private CodeAnalysisOptionsWrapper _codeAnalysisOptions;

        private bool _alwaysDisplayFileName;
        private bool _displayCRC32;
        private bool _displayContext;
        private bool _displayResults = true;
        private bool _displaySummary = true;
        private int _displaySummaryMinimum = 1;
        private bool _listImpactedFilesOnly;

        public CodeAnalysisOptionsWrapper CodeAnalysisOptions
        {
            get { return _codeAnalysisOptions; }
            set
            {
                if (_codeAnalysisOptions == value)
                    return;
                _codeAnalysisOptions = value;
                OnPropertyChanged();
            }
        }

        public bool AlwaysDisplayFileName
        {
            get { return _alwaysDisplayFileName; }
            set
            {
                if (_alwaysDisplayFileName == value)
                    return;
                _alwaysDisplayFileName = value;
                OnPropertyChanged();
            }
        }
  
        public bool DisplayContext
        {
            get { return _displayContext; }
            set
            {
                if (_displayContext == value)
                    return;
                _displayContext = value;
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
            get { return _displayCRC32; }
            set
            {
                if (_displayCRC32 == value)
                    return;
                _displayCRC32 = value;
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

        #endregion Fields and Properties

        #region Public Methods

        #endregion Public Methods

        #region Private Methods

        #endregion Private Methods

        #region IInstanceCount

        private static int _instanceCountVM;

        public int InstanceCountVM
        {
            get => _instanceCountVM;
            set => _instanceCountVM = value;
        }

        #endregion IInstanceCount
    }
}