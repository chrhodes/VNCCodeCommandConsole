using System;
using System.Collections.Generic;
using System.Windows;

using System.Xml.Linq;

using DevExpress.Xpf.Editors;

using VNC;
using VNC.Core.Mvvm;

using VNCCodeCommandConsole.Presentation.ViewModels;

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

        private void lbeContextMode_EditValueChanged(object sender, EditValueChangedEventArgs e)
        {
            var foo = e;
            var bar = (ListBoxEdit)sender;

            var barI = (ListBoxEditItem)bar.SelectedItem;

            switch (barI.Tag)
            {
                case "S":
                    lgContextDemo.IsCollapsed = true;
                    lgContextSolutionProject.IsCollapsed = false;
                    lgContextXmlConfig.IsCollapsed = true;
                    break;

                case "X":
                    lgContextDemo.IsCollapsed = true;
                    lgContextSolutionProject.IsCollapsed = true;
                    lgContextXmlConfig.IsCollapsed = false;
                    break;

                case "D":
                    lgContextDemo.IsCollapsed = false;
                    lgContextSolutionProject.IsCollapsed = true;
                    lgContextXmlConfig.IsCollapsed = true;
                    break;

                default:
                    throw new ArgumentException($"lbeContextMode: Unexpected tag ({bar.Tag})");
            }
        }

        private void RadioButton_Checked(object sender, RoutedEventArgs e)
        {
            var foo = e;
            var bar = (System.Windows.Controls.RadioButton)sender;

            switch (bar.Tag)
            {
                case "S":
                    lgContextDemo.IsCollapsed = true;
                    lgContextSolutionProject.IsCollapsed = false;
                    lgContextXmlConfig.IsCollapsed = true;
                    break;

                case "X":
                    lgContextDemo.IsCollapsed = true;
                    lgContextSolutionProject.IsCollapsed = true;
                    lgContextXmlConfig.IsCollapsed = false;
                    break;

                case "D":
                    lgContextDemo.IsCollapsed = false;
                    lgContextSolutionProject.IsCollapsed = true;
                    lgContextXmlConfig.IsCollapsed = true;
                    break;

                default:
                    throw new ArgumentException($"lbeContextMode: Unexpected tag ({bar.Tag})");
            }
        }

        private void lbeContextMode2_EditValueChanged(object sender, EditValueChangedEventArgs e)
        {
            var foo = e;
            var fooV = foo.NewValue;
            var bar = (ListBoxEdit)sender;

            var barI = bar.SelectedItem;
            var barT = barI.GetType();
            var bar2 = (DevExpress.Mvvm.EnumMemberInfo)barI;

            var foo2 = e.NewValue;
            var fooT = e.GetType();

            switch (e.NewValue.ToString())
            {
                case "SolutionProject":
                    lgContextDemo.IsCollapsed = true;
                    lgContextSolutionProject.IsCollapsed = false;
                    lgContextXmlConfig.IsCollapsed = true;
                    break;

                case "XmlConfig":
                    lgContextDemo.IsCollapsed = true;
                    lgContextSolutionProject.IsCollapsed = true;
                    lgContextXmlConfig.IsCollapsed = false;
                    break;

                case "Demo":
                    lgContextDemo.IsCollapsed = false;
                    lgContextSolutionProject.IsCollapsed = true;
                    lgContextXmlConfig.IsCollapsed = true;
                    break;

                default:
                    throw new ArgumentException($"lbeContextMode2: Unexpected value ({e.NewValue})");
            }

        }

        private void lbeContextMode3_EditValueChanged(object sender, EditValueChangedEventArgs e)
        {
            //var foo = e;
            //var fooV = foo.NewValue;
            //var bar = (ListBoxEdit)sender;

            //var barI = bar.SelectedItem;
            //var barT = barI.GetType();
            ////var bar2 = (DevExpress.Mvvm.EnumMemberInfo)barI;

            ////var foo2 = e.NewValue;
            ////var fooT = e.GetType();


            switch (e.NewValue.ToString())
            {
                case "SolutionProject":
                    lgContextSolutionProject.IsCollapsed = false;
                    lgContextXmlConfig.IsCollapsed = true;
                    lgContextFile.IsCollapsed = true;
                    lgContextDemo.IsCollapsed = true;
                    break;

                case "XmlConfig":
                    lgContextSolutionProject.IsCollapsed = true;
                    lgContextXmlConfig.IsCollapsed = false;
                    lgContextFile.IsCollapsed = true;
                    lgContextDemo.IsCollapsed = true;
                    break;

                case "File":
                    lgContextSolutionProject.IsCollapsed = true;
                    lgContextXmlConfig.IsCollapsed = true;
                    lgContextFile.IsCollapsed = false;
                    lgContextDemo.IsCollapsed = true;
                    break;

                case "Demo":
                    lgContextSolutionProject.IsCollapsed = true;
                    lgContextXmlConfig.IsCollapsed = true;
                    lgContextFile.IsCollapsed = true;
                    lgContextDemo.IsCollapsed = false;
                    break;

                default:
                    throw new ArgumentException($"lbeContextMode2: Unexpected value ({e.NewValue})");
            }
        }
    }
}
