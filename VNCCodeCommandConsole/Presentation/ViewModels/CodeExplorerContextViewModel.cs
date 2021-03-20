using System;
using System.Windows.Input;
using System.Xml;
using System.Xml.Linq;

using Prism.Commands;
using Prism.Events;

using VNC;
using VNC.Core.Mvvm;
using VNC.Core.Services;

using VNCCodeCommandConsole.Core.Events;

namespace VNCCodeCommandConsole.Presentation.ViewModels
{
    public class CodeExplorerContextViewModel : EventViewModelBase, ICodeExplorerContextViewModel, IInstanceCountVM
    {
        #region Constructors, Initialization, and Load

        public CodeExplorerContextViewModel(
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

            SayHelloCommand = new DelegateCommand(SayHello, SayHelloCanExecute);

            BrowseForFileCommand = new DelegateCommand(BrowseForFile, BrowseForFileCanExecute);
            ClearFileCommand = new DelegateCommand(ClearFile, ClearFileCanExecute);

            Message = "CodeExplorerContextViewModel says hello";

            PopulateContextFromXmlFile();

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void PopulateContextFromXmlFile()
        {
            long startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            Branches = new System.Collections.ObjectModel.ObservableCollection<XElement>();
            SolutionFiles = new System.Collections.ObjectModel.ObservableCollection<string>();
            ProjectFiles = new System.Collections.ObjectModel.ObservableCollection<string>();

            SolutionFiles = new System.Collections.ObjectModel.ObservableCollection<string>();

            XmlTextReader xtr = new XmlTextReader(Common.cCONFIG_FILE);

            XDocument xDocument = XDocument.Load(xtr, LoadOptions.PreserveWhitespace);

            var sourceBranches = xDocument.Descendants("SourceBranches");

            foreach (var branch in sourceBranches.Elements())
            {
                Branches.Add(branch);
                //Branches.Add(branch.Attribute("Name").Value);
            }

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion Constructors, Initialization, and Load

        #region Fields and Properties


        private string _selectedSourceFile;
        private string _selectedProject;
        private string _selectedSolution;
        private string _branch;
        private string _message;
        private string _projectFile;
        private string _repository;
        private string _repositoryPath;
        private string _solutionFile;
        private string _sourceFile;
        private string _sourcePath;

        public string Branch
        {
            get => _branch;
            set
            {
                if (_branch == value)
                    return;
                _branch = value;
                OnPropertyChanged();
            }
        }

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

        public string ProjectFile
        {
            get => _projectFile;
            set
            {
                if (_projectFile == value)
                    return;
                _projectFile = value;
                OnPropertyChanged();
            }
        }

        public string Repository
        {
            get => _repository;
            set
            {
                if (_repository == value)
                    return;
                _repository = value;
                OnPropertyChanged();
            }
        }

        public string RepositoryPath
        {
            get => _repositoryPath;
            set
            {
                if (_repositoryPath == value)
                    return;
                _repositoryPath = value;
                OnPropertyChanged();
            }
        }

        public ICommand SayHelloCommand { get; private set; }

        public string SolutionFile
        {
            get => _solutionFile;
            set
            {
                if (_solutionFile == value)
                    return;
                _solutionFile = value;
                OnPropertyChanged();
            }
        }

        public string SourceFile
        {
            get => _sourceFile;
            set
            {
                if (_sourceFile == value)
                    return;
                _sourceFile = value;
                OnPropertyChanged();
            }
        }

        public string SourcePath
        {
            get => _sourcePath;
            set
            {
                if (_sourcePath == value)
                    return;
                _sourcePath = value;
                OnPropertyChanged();
            }
        }

        public System.Collections.ObjectModel.ObservableCollection<XElement> Branches { get; set; }

        private XElement _selectedBranch;

        public XElement SelectedBranch
        {
            get => _selectedBranch;
            set
            {
                if (_selectedBranch == value)
                    return;
                _selectedBranch = value;

                SolutionFiles.Clear();

                var solutions = _selectedBranch.Descendants("Solution");

                foreach (var solution in solutions.Elements())
                {
                    SolutionFiles.Add(solution.Attribute("Name").Value);
                }

                OnPropertyChanged();
            }
        }
        
        public System.Collections.ObjectModel.ObservableCollection<string> SolutionFiles { get; set; }

        public string SelectedSolution
        {
            get => _selectedSolution;
            set
            {
                if (_selectedSolution == value)
                    return;
                _selectedSolution = value;
                OnPropertyChanged();
            }
        }

        public System.Collections.ObjectModel.ObservableCollection<string> ProjectFiles { get; set; }


        public string SelectedProject
        {
            get => _selectedProject;
            set
            {
                if (_selectedProject == value)
                    return;
                _selectedProject = value;
                OnPropertyChanged();
            }
        }

        public System.Collections.ObjectModel.ObservableCollection<string> SourceFiles { get; set; }

        
        public string SelectedSourceFile
        {
            get => _selectedSourceFile;
            set
            {
                if (_selectedSourceFile == value)
                    return;
                _selectedSourceFile = value;
                OnPropertyChanged();
            }
        }
        
        #endregion Fields and Properties

        #region Event Handlers

        #region BrowseForFile Command

        public DelegateCommand BrowseForFileCommand { get; set; }
        public string BrowseForFileContent { get; set; } = "BrowseForFile";
        public string BrowseForFileToolTip { get; set; } = "BrowseForFile ToolTip";

        // Can get fancy and use Resources
        //public string BrowseForFileContent { get; set; } = "ViewName_BrowseForFileContent";
        //public string BrowseForFileToolTip { get; set; } = "ViewName_BrowseForFileContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_BrowseForFileContent">BrowseForFile</system:String>
        //    <system:String x:Key="ViewName_BrowseForFileContentToolTip">BrowseForFile ToolTip</system:String>

        public void BrowseForFile()
        {
            // TODO(crhodes)
            // Do something amazing.
            Message = "Cool, you called BrowseForFile";
            EventAggregator.GetEvent<BrowseForFileEvent>().Publish();

            // Start Cut Four

            // Put this in places that listen for event
            //Common.EventAggregator.GetEvent<BrowseForFileEvent>().Subscribe(BrowseForFile);

            // End Cut Four
        }

        public bool BrowseForFileCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion BrowseForFile Command

        #region ClearFile Command

        public DelegateCommand ClearFileCommand { get; set; }
        public string ClearFileContent { get; set; } = "ClearFile";
        public string ClearFileToolTip { get; set; } = "ClearFile ToolTip";

        // Can get fancy and use Resources
        //public string ClearFileContent { get; set; } = "ViewName_ClearFileContent";
        //public string ClearFileToolTip { get; set; } = "ViewName_ClearFileContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_ClearFileContent">ClearFile</system:String>
        //    <system:String x:Key="ViewName_ClearFileContentToolTip">ClearFile ToolTip</system:String>

        public void ClearFile()
        {
            // TODO(crhodes)
            // Do something amazing.
            Message = "Cool, you called ClearFile";
            EventAggregator.GetEvent<ClearFileEvent>().Publish();

            // Start Cut Four

            // Put this in places that listen for event
            //Common.EventAggregator.GetEvent<ClearFileEvent>().Subscribe(ClearFile);

            // End Cut Four
        }

        public bool ClearFileCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion ClearFile Command

        #endregion Event Handlers

        #region Private Methods

        private void SayHello()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Message = "Hello";

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private bool SayHelloCanExecute()
        {
            return true;
        }

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