using System;
using System.Collections.Generic;
using System.Windows;
using System.Xml.Linq;

using DevExpress.Xpf.Editors;

using VNC;
using VNC.Core.Mvvm;

namespace VNCCodeCommandConsole.Presentation.Views
{
    public partial class CodeExplorerContext : ViewBase, ICodeExplorerContext, IInstanceCountV
    {

        public CodeExplorerContext()
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InstanceCountV++;
            InitializeComponent();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public CodeExplorerContext(ViewModels.ICodeExplorerContextViewModel viewModel)
        {
            Int64 startTicks = Log.CONSTRUCTOR($"Enter viewModel({viewModel.GetType()}", Common.LOG_CATEGORY);

            InstanceCountV++;
            InitializeComponent();

            ViewModel = viewModel;

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #region IInstanceCount

        private static int _instanceCountV;

        public int InstanceCountV
        {
            get => _instanceCountV;
            set => _instanceCountV = value;
        }

        #endregion


        private void cbeSolutionFile_EditValueChanged(object sender, DevExpress.Xpf.Editors.EditValueChangedEventArgs e)
        {
#if TRACE
            long startTicks = Log.PRESENTATION("Enter", Common.LOG_CATEGORY);
#endif
            try
            {
                var s = (ComboBoxEdit)sender;

                if (s.SelectedItems.Count == 1)
                {
                    XElement solution = (XElement)s.SelectedItem;

                    cbeSourceFiles.Items.Clear();
                    cbeSourceFiles.Clear();
                    cbeSourceFiles.ItemsSource = null;

                    cbeProjectFiles.Items.Clear();
                    cbeProjectFiles.Clear();
                    cbeProjectFiles.ItemsSource = solution.Elements("Project");

                    string fileName = solution.Attribute("FileName").Value;
                    string folderPath = solution.Attribute("FolderPath").Value;

                    //teSolutionFile.Text = teRepositoryPath.Text + "\\" + folderPath + "\\" + fileName;
                    //teSourcePath.Text = teRepositoryPath.Text + "\\" + folderPath + "\\";
                }
                else
                {
                    // Have selected multiple solution files, so clear out controls that affect GetFilesToProc()

                    List<XElement> projectElements = new List<XElement>();

                    foreach (XElement solutionElement in cbeSolutionFiles.SelectedItems)
                    {
                        foreach (XElement projectElement in solutionElement.Elements("Project"))
                        {
                            projectElements.Add(projectElement);
                        }
                    }

                    cbeProjectFiles.Items.Clear();
                    cbeProjectFiles.Clear();
                    cbeProjectFiles.ItemsSource = projectElements;
                    //teProjectFile.Clear();

                    cbeSourceFiles.Items.Clear();
                    cbeSourceFiles.Clear();
                    cbeSourceFiles.ItemsSource = null;
                    //teSourceFile.Clear();

                    //teSourcePath.Clear();   // not sure if this should be cleared.

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
#if TRACE
            Log.PRESENTATION("End", Common.LOG_CATEGORY, startTicks);
#endif
        }
    }
}
