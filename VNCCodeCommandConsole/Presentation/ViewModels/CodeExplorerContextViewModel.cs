using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
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

            Branches = new ObservableCollection<XElement>();
            SolutionFiles = new ObservableCollection<XElement>();
            ProjectFiles = new ObservableCollection<XElement>();
            SourceFiles = new ObservableCollection<XElement>();

            SelectedSolutionFiles = new ObservableCollection<XElement>();
            SelectedProjectFiles = new ObservableCollection<XElement>();
            SelectedSourceFiles = new ObservableCollection<XElement>();

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

        public ObservableCollection<XElement> Branches { get; set; }

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
                ProjectFiles.Clear();
                SourceFiles.Clear();

                var solutions = _selectedBranch.Descendants("Solution");

                SolutionFiles.AddRange(solutions);

                OnPropertyChanged();
            }
        }
        
        public ObservableCollection<XElement> SolutionFiles { get; set; }

        ObservableCollection<XElement> _selectedSolutionFiles;

        public ObservableCollection<XElement> SelectedSolutionFiles
        { 
            get => _selectedSolutionFiles;
            set
            {
                if (_selectedSolutionFiles == value)
                    return;
                _selectedSolutionFiles = value;

                ProjectFiles.Clear();

                var projects = _selectedSolutionFiles.Descendants("Project");
                ProjectFiles.AddRange(projects);

                OnPropertyChanged();
            }
        }

        private XElement _selectedSolution;

        public XElement SelectedSolution
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

        public ObservableCollection<XElement> ProjectFiles { get; set; }

        ObservableCollection<XElement> _selectedProjectFiles;

        public ObservableCollection<XElement> SelectedProjectFiles
        {
            get => _selectedProjectFiles;
            set
            {
                if (_selectedProjectFiles == value)
                    return;
                _selectedProjectFiles = value;

                // TODO(crhodes)
                // What now, get list of Source Files?
                var sourceFiles = _selectedSolutionFiles.Descendants("SourceFile");
                SourceFiles.AddRange(sourceFiles);

                OnPropertyChanged();
            }
        }

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

        public ObservableCollection<XElement> SourceFiles { get; set; }

        ObservableCollection<XElement> _selectedSourceFiles;

        public ObservableCollection<XElement> SelectedSourceFiles
        {
            get => _selectedSourceFiles;
            set
            {
                if (_selectedSourceFiles == value)
                    return;
                _selectedSourceFiles = value;

                OnPropertyChanged();
            }
        }

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

        public string GetFilePath(XElement sourceFileElement)
        {
            string fileName = sourceFileElement.Attribute("FileName").Value;
            string folderPath = sourceFileElement.Attribute("FolderPath").Value;
            string filePath = (folderPath != "" ? folderPath + "\\" : "") + fileName;

            return filePath;
        }

        // TODO(crhodes)
        // This should be calling a DOMAIN/APPLICATION Service

        /// <summary>
        /// Returns list of files to process based on selections
        /// </summary>
        /// <returns></returns>
        public List<String> GetFilesToProcess()
        {
            long startTicks = Log.PRESENTATION("Enter", Common.LOG_CATEGORY);

            List<String> filesToProcess = new List<string>();

            // HACK(crhodes)
            // Just hard code a couple of files for now till we sort through this

            filesToProcess.Add(@"C:\Temp\VNCCodeCommandConsoleTestFiles\AML.vb");
            filesToProcess.Add(@"C:\Temp\VNCCodeCommandConsoleTestFiles\AppConfig.vb");

            filesToProcess.Add(@"C:\Temp\VNCCodeCommandConsoleTestFiles\modEASE1.vb");
            filesToProcess.Add(@"C:\Temp\VNCCodeCommandConsoleTestFiles\modEASE2.vb");
            filesToProcess.Add(@"C:\Temp\VNCCodeCommandConsoleTestFiles\modEASE3.vb");
            filesToProcess.Add(@"C:\Temp\VNCCodeCommandConsoleTestFiles\modEASE4.vb");

            filesToProcess.Add(@"C:\Temp\VNCCodeCommandConsoleTestFiles\VB.cs");
            filesToProcess.Add(@"C:\Temp\VNCCodeCommandConsoleTestFiles\CS.cs");
            filesToProcess.Add(@"C:\Temp\VNCCodeCommandConsoleTestFiles\VNCCSSyntaxWalkerBase.cs");
            filesToProcess.Add(@"C:\Temp\VNCCodeCommandConsoleTestFiles\VNCVBSyntaxWalkerBase.cs");

            //// This  method returns a list of files to process.
            //// If a specific SourceFile is specified, return it
            ////
            //// If multiple SolutionFiles are selected, the user must select
            ////  one or more project files.  Handle as multiple project files
            ////
            //// If multiple ProjectFiles are selected, 
            ////  Loop across <Project> elements 
            ////      and add files from project
            ////      and add files listed in <Project> elements
            ////
            //// If a ProjectFile is available, use it to get the list of files
            //// Otherwise return the files selected in cbeSourceFiles.

            //string solutionFullPath = SolutionFile;
            //string projectFullPath = ProjectFile;

            //if (SourceFile != "")
            //{
            //    // TODO(crhodes)
            //    // Add check for existence
            //    filesToProcess.Add(SourceFile);
            //}
            //else if (SelectedProjectFiles.Count > 1)
            //{
            //    foreach (XElement projectElement in SelectedProjectFiles)
            //    {
            //        string fileName = projectElement.Attribute("FileName").Value;
            //        string folderPath = projectElement.Attribute("FolderPath").Value;
            //        string sourcePath = RepositoryPath + "\\" + folderPath;
            //        string projectPath = "";

            //        projectPath = fileName != "" ? sourcePath + "\\" + fileName : "";

            //        if (projectPath == "")
            //        {
            //            // No project file exists, so look across all the SourceFile elements

            //            foreach (XElement sourceFile in projectElement.Elements("SourceFile"))
            //            {
            //                // NB. The file names are added manually so we don't have to exclude any.
            //                string fileFullPath = sourcePath + "\\" + GetFilePath(sourceFile);

            //                filesToProcess.Add(fileFullPath);
            //            }
            //        }
            //        else
            //        {
            //            filesToProcess.AddRange(VNC.CodeAnalysis.Workspace.Helper.GetSourceFilesToProcessFromVSProject(projectPath));
            //        }
            //    }
            //}
            //else if (projectFullPath != "")
            //{
            //    filesToProcess = VNC.CodeAnalysis.Workspace.Helper.GetSourceFilesToProcessFromVSProject(projectFullPath);
            //}
            //else if (SelectedSourceFiles.Count > 0)
            //{
            //    string sourcePath = SourcePath;

            //    foreach (XElement sourceFile in SelectedSourceFiles)
            //    {
            //        string fileFullPath = sourcePath + "\\" + GetFilePath(sourceFile);

            //        filesToProcess.Add(fileFullPath);
            //    }
            //}

            //var filesToCheck = filesToProcess.ToList();

            //foreach (string filePath in filesToCheck)
            //{
            //    if (!File.Exists(filePath))
            //    {
            //        FileInfo fileInfo = new FileInfo(filePath);

            //        if (!Directory.Exists(fileInfo.DirectoryName))
            //        {
            //            MessageBox.Show(string.Format("Directory\n\n{0}\n\ndoes not exist", fileInfo.DirectoryName), "Check Path or Config File");
            //        }
            //        else
            //        {
            //            MessageBox.Show(string.Format("File\n\n{0}\nin\n\n{1}\n\ndoes not exist", fileInfo.Name, fileInfo.DirectoryName), "Check Path or Config File");
            //        }

            //        filesToProcess.Remove(filePath);
            //    }
            //}

            Log.PRESENTATION($"End: filesToProcess.Count {filesToProcess.Count}", Common.LOG_CATEGORY, startTicks);

            return filesToProcess;
        }
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