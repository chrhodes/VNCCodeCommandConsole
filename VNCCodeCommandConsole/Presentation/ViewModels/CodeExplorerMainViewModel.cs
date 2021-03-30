using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.VisualBasic;

using Prism.Events;

using VNC;
using VNC.Core.Mvvm;
using VNC.Core.Services;

using VNCCodeCommandConsole.Core.Events;

using VNCCA = VNC.CodeAnalysis;
using VNCSW = VNC.CodeAnalysis.SyntaxWalkers;

namespace VNCCodeCommandConsole.Presentation.ViewModels
{
    public class CodeExplorerMainViewModel : EventViewModelBase, ICodeExplorerMainViewModel, IInstanceCountVM
    {
        private string _workspace;
        private string _syntaxStructuredTrivia;
        private string _syntaxTrivia;
        private string _syntaxToken;
        private string _syntaxNode;
        private string _syntaxTree;
        private string _summaryCRCToFullString;
        private string _summaryCRCToString;
        private string _sourceCode = "Source Code Output Goes Here";
        private string _title = "VNCCodeCommandConsole - MainWindowDxDockLayoutManager";


        private ConfigurationOptionsViewModel _configurationOptionsViewModel;
        private CodeExplorerContextViewModel _codeExplorerContextViewModel;

        public string Title
        {
            get => _title;
            set
            {
                if (_title == value)
                    return;
                _title = value;
                OnPropertyChanged();
            }
        }

        // ObservableCollection notifies databinding when collection changes
        // because it implements INotifyCollectionChanged

        //public ObservableCollection<Address> Addresses { get; set; }

        //public ObservableCollection<Material> Materials { get; set; }

        //public IAddressDataService _addressDataService { get; set; }

        //public IMaterialDataService _materialDataService { get; set; }
        public CodeExplorerMainViewModel(
            IConfigurationOptionsViewModel configurationOptionsViewModel,
            CodeExplorerContextViewModel codeExplorerContextViewModel,
            IEventAggregator eventAggregator,
            IMessageDialogService messageDialogService) : base(eventAggregator, messageDialogService)
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            // HACK(crhodes)
            // Should add what we need to IConfigurationOptionsViewModel

            _configurationOptionsViewModel = (ConfigurationOptionsViewModel)configurationOptionsViewModel;
            _codeExplorerContextViewModel = codeExplorerContextViewModel;
            InitializeViewModel();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void InitializeViewModel()
        {
            Int64 startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            InstanceCountVM++;

            EventAggregator.GetEvent<SyntaxWalkerResultEvent>().Subscribe(DisplayResults);
            EventAggregator.GetEvent<InvokeCSSyntaxWalkerEvent>().Subscribe(ProcessOperationCS);
            EventAggregator.GetEvent<InvokeVBSyntaxWalkerEvent>().Subscribe(ProcessOperationVB);

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void DisplayResults(string obj)
        {
            Summary = obj;
        }

        public void ProcessOperationVB(VNCCA.Types.SearchTreeCommand searchTreeCommand)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            StringBuilder sb = new StringBuilder();

            ClearPreviousResults();

            Dictionary<string, Int32> matches = new Dictionary<string, int>();
            Dictionary<string, Int32> crcMatchesToString = new Dictionary<string, int>();
            Dictionary<string, Int32> crcMatchesToFullString = new Dictionary<string, int>();

            var filesToProcess = _codeExplorerContextViewModel.GetFilesToProcess();

            if (filesToProcess.Count > 0)
            {
                if ((Boolean)_configurationOptionsViewModel.ListImpactedFilesOnly)
                {
                    sb.AppendLine("Would Search these files ....");
                }

                foreach (string filePath in filesToProcess)
                {
                    if ((Boolean)_configurationOptionsViewModel.ListImpactedFilesOnly)
                    {
                        sb.AppendLine(string.Format("  {0}", filePath));
                    }
                    else
                    {
                        StringBuilder sbFileResults = new StringBuilder();

                        var sourceCode = "";

                        using (var sr = new System.IO.StreamReader(filePath))
                        {
                            sourceCode = sr.ReadToEnd();
                        }

                        // 
                        // This is where the action happens
                        //

                        SyntaxTree tree = VisualBasicSyntaxTree.ParseText(sourceCode);

                        VNCCA.SearchTreeCommandConfiguration searchTreeCommandConfiguration = new VNCCA.SearchTreeCommandConfiguration();

                        searchTreeCommandConfiguration.CodeAnalysisOptions = _configurationOptionsViewModel.CodeAnalysisOptions.Model;

                        searchTreeCommandConfiguration.Results = sbFileResults;
                        searchTreeCommandConfiguration.SyntaxTree = tree;
                        searchTreeCommandConfiguration.Matches = matches;
                        searchTreeCommandConfiguration.CRCMatchesToString = crcMatchesToString;
                        searchTreeCommandConfiguration.CRCMatchesToFullString = crcMatchesToFullString;

                        sbFileResults = searchTreeCommand(searchTreeCommandConfiguration);

                        if ((bool)_configurationOptionsViewModel.AlwaysDisplayFileName || (sbFileResults.Length > 0))
                        {
                            sb.AppendLine("Searching " + filePath);
                        }

                        sb.Append(sbFileResults.ToString());
                    }
                }
            }
            else
            {
                sb.AppendLine("No files selected to process");
            }

            if (!(Boolean)_configurationOptionsViewModel.DisplayResults)
            {
                // If only want to see the summary ...
                sb.Clear();
            }

            SourceCode = sb.ToString();

            if ((Boolean)_configurationOptionsViewModel.DisplaySummary)
            {
                StringBuilder summary = new StringBuilder();

                // Add information from the matches dictionary
                summary.AppendLine("\n*** Summary ***\n");

                foreach (var item in matches.OrderBy(v => v.Key).Select(k => k.Key))
                {
                    if (matches[item] >= _configurationOptionsViewModel.DisplaySummaryMinimum)
                    {
                        summary.AppendLine(string.Format("Count: {0,3} {1} ", matches[item], item));
                    }
                }

                Summary = summary.ToString();

                if ((Boolean)_configurationOptionsViewModel.DisplayCRC32)
                {
                    summary.Clear();

                    summary.AppendLine("\n*** CRC ToString Summary *** \n");

                    foreach (var item in crcMatchesToString.OrderBy(v => v.Key).Select(k => k.Key))
                    {
                        if (crcMatchesToString[item] >= _configurationOptionsViewModel.DisplaySummaryMinimum)
                        {
                            summary.AppendLine(string.Format("Count: {0,3} {1} ", crcMatchesToString[item], item));
                        }
                    }

                    SummaryCRCToString = summary.ToString();

                    summary.Clear();

                    summary.AppendLine("\n*** CRC ToFullString Summary ***\n");

                    foreach (var item in crcMatchesToFullString.OrderBy(v => v.Key).Select(k => k.Key))
                    {
                        if (crcMatchesToFullString[item] >= _configurationOptionsViewModel.DisplaySummaryMinimum)
                        {
                            summary.AppendLine(string.Format("Count: {0,3} {1} ", crcMatchesToFullString[item], item));
                        }
                    }

                    SummaryCRCToString = summary.ToString();
                }

            }

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public void ProcessOperationCS(VNCCA.Types.SearchTreeCommand searchTreeCommand)
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            StringBuilder sb = new StringBuilder();

            ClearPreviousResults();

            Dictionary<string, Int32> matches = new Dictionary<string, int>();
            Dictionary<string, Int32> crcMatchesToString = new Dictionary<string, int>();
            Dictionary<string, Int32> crcMatchesToFullString = new Dictionary<string, int>();

            var filesToProcess = _codeExplorerContextViewModel.GetFilesToProcess();

            if (filesToProcess.Count > 0)
            {
                if ((Boolean)_configurationOptionsViewModel.ListImpactedFilesOnly)
                {
                    sb.AppendLine("Would Search these files ....");
                }

                foreach (string filePath in filesToProcess)
                {
                    if ((Boolean)_configurationOptionsViewModel.ListImpactedFilesOnly)
                    {
                        sb.AppendLine(string.Format("  {0}", filePath));
                    }
                    else
                    {
                        StringBuilder sbFileResults = new StringBuilder();

                        var sourceCode = "";

                        using (var sr = new System.IO.StreamReader(filePath))
                        {
                            sourceCode = sr.ReadToEnd();
                        }

                        // 
                        // This is where the action happens
                        //

                        SyntaxTree tree = CSharpSyntaxTree.ParseText(sourceCode);

                        VNCCA.SearchTreeCommandConfiguration searchTreeCommandConfiguration = new VNCCA.SearchTreeCommandConfiguration();

                        searchTreeCommandConfiguration.CodeAnalysisOptions = _configurationOptionsViewModel.CodeAnalysisOptions.Model;

                        searchTreeCommandConfiguration.Results = sbFileResults;
                        searchTreeCommandConfiguration.SyntaxTree = tree;
                        searchTreeCommandConfiguration.Matches = matches;
                        searchTreeCommandConfiguration.CRCMatchesToString = crcMatchesToString;
                        searchTreeCommandConfiguration.CRCMatchesToFullString = crcMatchesToFullString;

                        sbFileResults = searchTreeCommand(searchTreeCommandConfiguration);

                        if ((bool)_configurationOptionsViewModel.AlwaysDisplayFileName || (sbFileResults.Length > 0))
                        {
                            sb.AppendLine("Searching " + filePath);
                        }

                        sb.Append(sbFileResults.ToString());
                    }
                }
            }
            else
            {
                sb.AppendLine("No files selected to process");
            }

            if (!(Boolean)_configurationOptionsViewModel.DisplayResults)
            {
                // If only want to see the summary ...
                sb.Clear();
            }

            SourceCode = sb.ToString();

            if ((Boolean)_configurationOptionsViewModel.DisplaySummary)
            {
                StringBuilder summary = new StringBuilder();

                // Add information from the matches dictionary
                summary.AppendLine("\n*** Summary ***\n");

                foreach (var item in matches.OrderBy(v => v.Key).Select(k => k.Key))
                {
                    if (matches[item] >= _configurationOptionsViewModel.DisplaySummaryMinimum)
                    {
                        summary.AppendLine(string.Format("Count: {0,3} {1} ", matches[item], item));
                    }
                }

                Summary = summary.ToString();

                if ((Boolean)_configurationOptionsViewModel.DisplayCRC32)
                {
                    summary.Clear();

                    summary.AppendLine("\n*** CRC ToString Summary *** \n");

                    foreach (var item in crcMatchesToString.OrderBy(v => v.Key).Select(k => k.Key))
                    {
                        if (crcMatchesToString[item] >= _configurationOptionsViewModel.DisplaySummaryMinimum)
                        {
                            summary.AppendLine(string.Format("Count: {0,3} {1} ", crcMatchesToString[item], item));
                        }
                    }

                    SummaryCRCToString = summary.ToString();

                    summary.Clear();

                    summary.AppendLine("\n*** CRC ToFullString Summary ***\n");

                    foreach (var item in crcMatchesToFullString.OrderBy(v => v.Key).Select(k => k.Key))
                    {
                        if (crcMatchesToFullString[item] >= _configurationOptionsViewModel.DisplaySummaryMinimum)
                        {
                            summary.AppendLine(string.Format("Count: {0,3} {1} ", crcMatchesToFullString[item], item));
                        }
                    }

                    SummaryCRCToString = summary.ToString();
                }

            }

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void ClearPreviousResults()
        {
            // TODO(crhodes)
            // Create Properties that View can bind to

            SourceCode = "";
            SyntaxNode = "";
            SyntaxToken = "";
            SyntaxTrivia = "";
            SyntaxStructuredTrivia = "";
            Summary = "";
            SummaryCRCToString = "";
            SummaryCRCToFullString = "";

            //teSourceCode.Clear();
            //teSourceCode.InvalidateVisual();

            //teSyntaxNode.Clear();
            //teSyntaxToken.Clear();
            //teSyntaxTrivia.Clear();
            //teSyntaxStructuredTrivia.Clear();

            //teSummary.Clear();
            //teSummaryCRCToString.Clear();
            //teSummaryCRCToFullString.Clear();
        }

        private string _summary = "TBD";
        public string Summary
        {
            get => _summary;
            set
            {
                if (_summary == value)
                    return;
                _summary = value;
                OnPropertyChanged();
            }
        }

        public string SourceCode
        {
            get => _sourceCode;
            set
            {
                if (_sourceCode == value)
                    return;
                _sourceCode = value;
                OnPropertyChanged();
            }
        }


        public string SummaryCRCToString
        {
            get => _summaryCRCToString;
            set
            {
                if (_summaryCRCToString == value)
                    return;
                _summaryCRCToString = value;
                OnPropertyChanged();
            }
        }

        public string SummaryCRCToFullString
        {
            get => _summaryCRCToFullString;
            set
            {
                if (_summaryCRCToFullString == value)
                    return;
                _summaryCRCToFullString = value;
                OnPropertyChanged();
            }
        }

        public string SyntaxTree
        {
            get => _syntaxTree;
            set
            {
                if (_syntaxTree == value)
                    return;
                _syntaxTree = value;
                OnPropertyChanged();
            }
        }


        public string SyntaxNode
        {
            get => _syntaxNode;
            set
            {
                if (_syntaxNode == value)
                    return;
                _syntaxNode = value;
                OnPropertyChanged();
            }
        }


        public string SyntaxToken
        {
            get => _syntaxToken;
            set
            {
                if (_syntaxToken == value)
                    return;
                _syntaxToken = value;
                OnPropertyChanged();
            }
        }


        public string SyntaxTrivia
        {
            get => _syntaxTrivia;
            set
            {
                if (_syntaxTrivia == value)
                    return;
                _syntaxTrivia = value;
                OnPropertyChanged();
            }
        }


        public string SyntaxStructuredTrivia
        {
            get => _syntaxStructuredTrivia;
            set
            {
                if (_syntaxStructuredTrivia == value)
                    return;
                _syntaxStructuredTrivia = value;
                OnPropertyChanged();
            }
        }
        
        public string DisplaySummary
        {
            get => _workspace;
            set
            {
                if (_workspace == value)
                    return;
                _workspace = value;
                OnPropertyChanged();
            }
        }
        


        //public MainWindowDxViewModel(
        //            IAddressDataService addressDataService,
        //            IMaterialDataService materialDataService)
        //{
        //    _addressDataService = addressDataService;
        //    _materialDataService = materialDataService;

        //    Addresses = new ObservableCollection<Address>();
        //    Materials = new ObservableCollection<Material>();
        //}

        #region This goes away with the new NavigationViewModel


        //Address _selectedAddress;

        //public Address SelectedAddress
        //{
        //    get { return _selectedAddress; }
        //    set
        //    {
        //        _selectedAddress = value;

        //        RaisePropertyChanged();
        //        // TODO(crhodes)
        //        // Learn what SetProperty does
        //        //SetProperty(ref _selectedCustomer, value);
        //    }
        //}

        //Material _selectedMaterial;

        //public Material SelectedMaterial
        //{
        //    get { return _selectedMaterial; }
        //    set
        //    {
        //        _selectedMaterial = value;

        //        RaisePropertyChanged();
        //        SetProperty(ref _selectedMaterial, value);
        //    }
        //}

        #endregion

        #region Move to Async loading

        //public void Load()
        //{
        //    LoadAddresses();
        //    LoadMaterials();
        //}

        //public async Task LoadAsync()
        //{
        //    await LoadAddressesAsync();
        //    await LoadMaterialsAsync();
        //}

        //private void LoadAddresses()
        //{
        //    var addresses = _addressDataService.All();

        //    Addresses.Clear();

        //    foreach (var address in addresses)
        //    {
        //        Addresses.Add(address);
        //    }
        //}

        //public async Task LoadAddressesAsync()
        //{
        //    var addresses = await _addressDataService.AllAsync();

        //    Addresses.Clear();

        //    foreach (var address in addresses)
        //    {
        //        Addresses.Add(address);
        //    }
        //}

        //private void LoadMaterials()
        //{
        //    var materials = _materialDataService.All();

        //    Materials.Clear();

        //    foreach (var material in materials)
        //    {
        //        Materials.Add(material);
        //    }
        //}

        //public async Task LoadMaterialsAsync()
        //{
        //    var materials = await _materialDataService.AllAsync();

        //    Materials.Clear();

        //    foreach (var material in materials)
        //    {
        //        Materials.Add(material);
        //    }
        //}

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
