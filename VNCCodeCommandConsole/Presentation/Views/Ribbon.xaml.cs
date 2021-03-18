using System;
using System.Windows;
using System.Windows.Controls;

using DevExpress.Xpf.Bars;
using DevExpress.Xpf.Core;
using DevExpress.Xpf.Grid;
using DevExpress.Xpf.Ribbon;


using VNC;
using VNC.Core.Mvvm;

namespace VNCCodeCommandConsole.Presentation.Views
{
    public partial class Ribbon : VNC.Core.Mvvm.ViewBase, IRibbon, IInstanceCountV
    {

        public Ribbon(ViewModels.IRibbonViewModel viewModel)
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

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

        #region Stuff from Code Behind.  Move to ViewModel


        #region Initialization

        //private void OnWindowLoaded(object sender, RoutedEventArgs e)
        //        {
        //#if TRACE
        //            long startTicks = VNC.Log.Trace15("Enter", Common.LOG_CATEGORY, CLASS_BASE_ERRORNUMBER + 0);
        //#endif
        //            // Do not load your data at design time.
        //            if (!System.ComponentModel.DesignerProperties.GetIsInDesignMode(this))
        //            {
        //                if (!Data.Config.DBBypass)
        //                {
        //                    Common.ApplicationDataSet.LoadApplicationDataSetFromDB(Common.ApplicationDataSet);
        //                }

        //                // HACK(crhodes)
        //                // Disable loading data so we can run app
        //                // but pretend we have
        //                Common.DataFullyLoaded = true;
        //            }

        //            VNC.AssemblyHelper.AssemblyInformation info = new VNC.AssemblyHelper.AssemblyInformation(System.Reflection.Assembly.GetExecutingAssembly());

        //            var eventMessage = string.Format("Started Version: {0}", info.InformationalVersionAttribute);

        //            Common.IndicateApplicationUsage(Common.LOG_CATEGORY, DateTime.Now, Common.CurrentUser.Identity.Name, eventMessage);

        //            Common.UserMode = new ViewMode(Data.Config.DefaultUserMode);
        //            Common.RowDetailMode = Data.Config.RowDetailMode;

        //#if TRACE
        //            VNC.Log.Trace15("Exit", Common.LOG_CATEGORY, CLASS_BASE_ERRORNUMBER + 2, startTicks);
        //#endif
        //}

        #endregion

        #region Event Handlers

        private void bei_DetailViewMode_EditValueChanged(object sender, RoutedEventArgs e)
        {
            //var bar = (BarEditItem)sender;

            //// TODO(crhodes): How to get to the tag value.  This keeps returning null.
            ////Common.RowDetailMode = (string)bar.Tag;

            //// For now use the edit value.  Need to update Switch in SetRowDetailType, too.
            //Common.RowDetailMode = bar.EditValue.ToString();

            //TableView tableView = FindTableView();

            //if (tableView != null)
            //{
            //    SetRowDetailType(tableView);
            //}
        }

        //private void HookTitleEvent(wucDXBase control)
        //{
        //    SetTitle(control, EventArgs.Empty);

        //    if (control != null)
        //    {
        //        control.TitleChanged += SetTitle;
        //        //control.TitleChanged2 += SetTitle2;
        //    }
        //}

        private void OnAboutClick(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e)
        {
            //try
            //{
            //    var win1 = new About();
            //    win1.Show();
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.ToString());
            //    MessageBox.Show(ex.InnerException.ToString());
            //}
        }

        private void OnAdministratorMode_Click(object sender, ItemClickEventArgs e)
        {
            //var snd = (BarCheckItem)sender;

            //if (snd.IsChecked == true)
            //{
            //    Common.UserMode.CurrentMode = (int)ViewMode.Mode.Administrator;
            //    Common.AllowEditing = true;
            //    lc_Root.Background = new SolidColorBrush(Color.FromArgb(0x7E, 0xFF, 0x01, 0x01));
            //}
            //else
            //{
            //    // TODO(crhodes): Decide if this is correct or if we need to use the Common.UserMode at all.
            //    Common.UserMode.CurrentMode = (int)ViewMode.Mode.Basic;
            //    Common.AllowEditing = false;
            //    lc_Root.Background = Brushes.Transparent;
            //}

            //RefreshScreen();
        }

        private void OnBetaMode_Clicked(object sender, ItemClickEventArgs e)
        {
            //var snd = (BarCheckItem)sender;

            //if (snd.IsChecked == true)
            //{
            //    Common.UserMode.CurrentMode = (int)ViewMode.Mode.Beta;
            //}
            //else
            //{
            //    Common.UserMode.CurrentMode = (int)ViewMode.Mode.Basic;
            //}

            //// This saves the current mode back to the app.config file.
            //Data.Config.DefaultUserMode = Common.UserMode.CurrentMode;

            //RefreshScreen();

            ////if (_currentControl != null)
            ////{
            ////    //UpdateAvailableViews();
            ////    // Redisplay the current control to allow updates to behavior based on mode.
            ////    // This may just affect the Menu/Ribbon in the future.
            ////    var currentType = _currentControl.GetType();

            ////    DisplayScreen(currentType.FullName);
            ////}
        }

        private void OnCheckItem(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e)
        {
            MessageBox.Show("OnCheckItem()");
        }

        private void OnDeveloperMode_Click(object sender, ItemClickEventArgs e)
        {
            //var snd = (BarCheckItem)sender;

            //if (snd.IsChecked == true)
            //{
            //    Common.DeveloperMode = true;
            //    //lc_Root.Background = new SolidColorBrush(Color.FromArgb(0x7E, 0xFF, 0x01, 0x01));
            //}
            //else
            //{
            //    // TODO(crhodes): Decide if this is correct or if we need to use the Common.UserMode at all.
            //    Common.DeveloperMode = false;
            //    lc_Root.Background = Brushes.Transparent;
            //}

            //RefreshScreen();
        }

        private void OnEditValueChanged_FirstSplash(object sender, DevExpress.Xpf.Editors.EditValueChangedEventArgs e)
        {
            //var ev = e;
            //var foo = e.NewValue;

            //string typeName = string.Format("{0}", ((SplashItem)e.NewValue).Value);

            //DisplayScreen(typeName);
        }

        private void OnEditValueChanged_ViewMode(object sender, DevExpress.Xpf.Editors.EditValueChangedEventArgs e)
        {
            var userMode = e.NewValue;
        }

        private void RefreshScreen()
        {
            //if (_currentControl != null)
            //{

            //    // Redisplay the current control to allow updates to behavior based on mode.
            //    // This may just affect the Menu/Ribbon in the future.
            //    var currentType = _currentControl.GetType();

            //    DisplayScreen(currentType.AssemblyQualifiedName);
            //    //DisplayScreen(currentType.FullName);
            //}
        }

        private void OnAdvancedMode_Click(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e)
        {
            //var snd = (BarCheckItem)sender;

            //if (snd.IsChecked == true)
            //{
            //    Common.UserMode.CurrentMode = (int)ViewMode.Mode.Advanced;
            //    IsAdvancedMode = true;

            //}
            //else
            //{
            //    Common.UserMode.CurrentMode = (int)ViewMode.Mode.Basic;
            //    IsAdvancedMode = false;
            //}

            //// This saves the current mode back to the app.config file.
            //Data.Config.DefaultUserMode = Common.UserMode.CurrentMode;

            ////UpdateAvailableViews();

            //RefreshScreen();
        }


        private void OnExit_Click(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e)
        {
            MessageBox.Show("OnExitClick()");
        }

        private void OnFileSave_Click(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e)
        {
            string controlTag = (string)e.Item.Tag;
            var tableView = FindTableView();

            if (tableView != null)
            {
                switch (controlTag)
                {

                    case "csv":
                        SaveDataToCSV(tableView);
                        break;

                    case "xls":
                        SaveDataToXLS(tableView);
                        break;

                    case "xlsx":
                        SaveDataToXLSX(tableView);
                        break;

                    default:
                        break;
                }
            }
        }

        private void OnGetHelpOn(object sender, RoutedEventArgs e)
        {
            //string helpTopic = ((Button)sender).Tag.ToString();

            //Helpers.Help.GetHelpOnTopic(helpTopic);
        }

        private void OnHelp_Click(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e)
        {
            //string helpTopic = ((Button)sender).Tag.ToString();

            //Helpers.Help.GetHelpOnTopic(helpTopic);
        }

        private void OnSendFeedback_Click(object sender, RoutedEventArgs e)
        {
            //ExternalProgram.SendFeedback();
        }

        private void OnSendFeedback_Click(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e)
        {
            //ExternalProgram.SendFeedback();
        }

        private void OnSettings_Click(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e)
        {
            //var window = new User_Interface.Windows.GlobalSettings();
            //window.ShowDialog();
        }

        private void OnShowInCurrentWindow_Click(object sender, ItemClickEventArgs e)
        {
            //string controlTag = (string)e.Item.Tag;

            //string typeName = string.Format("{0}", controlTag);
            ////string typeName = string.Format("VNCCodeCommandConsole.User_Interface.{0}", controlTag);

            ////if (Keyboard.IsKeyDown(Key.LeftCtrl))
            //if (Keyboard.Modifiers == ModifierKeys.Control)
            //{
            //    var win = new UserControlHost(typeName, typeName);
            //    win.Show();
            //}
            //else
            //{
            //    DisplayScreen(typeName);
            //}
        }

        private void OnShowInNewWindow_Click(object sender, ItemClickEventArgs e)
        {
            //string controlTag = (string)e.Item.Tag;

            //string typeName = string.Format("{0}", controlTag);
            ////string typeName = string.Format("VNCCodeCommandConsole.User_Interface.{0}", controlTag);

            ////if (Keyboard.IsKeyDown(Key.LeftCtrl))
            //if (Keyboard.Modifiers == ModifierKeys.Control)
            //{
            //    DisplayScreen(typeName);
            //}
            //else
            //{
            //    var win = new UserControlHost(typeName.Split('.').Last(), typeName);
            //    win.Show();
            //}
        }

        #endregion

        #region Main Function Routines

        private void DisplayScreen(string typeName)
        {
            ////var loadedAssemblies = AppDomain.CurrentDomain.GetAssemblies();

            ////foreach (var item in loadedAssemblies)
            ////{
            ////    AppLog.Info(item.FullName, LOG_CATEGORY);
            ////}

            ////System.Reflection.Assembly vncAssembly = System.Reflection.Assembly.GetAssembly(Type.GetType("VNC.WPFUserControls.wucUIOne"));
            //try
            //{

            //    //Type vnc = typeof(VNC.WPFUserControls.wucUIOne);
            //    //System.Reflection.Assembly vncAssembly = System.Reflection.Assembly.GetAssembly(vnc);

            //    Type ucType = Type.GetType(typeName);

            //    //var uc1 = Activator.CreateInstance(vnc);


            //    var uc = Activator.CreateInstance(ucType);
            //    //Type ty = uc.GetType()
            //    //ShowUserControl((UserControl)uc1);
            //    ShowUserControl((UserControl)uc);
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(
            //        string.Format("Cannot load type:{0}.  Check Tag Name.\n {1}",
            //        typeName,
            //        ex));
            //}
        }

        private GridControl FindGridControl()
        {
            //if (_currentControl != null)
            //{
            //    var gc = _currentControl.FindName("dataGrid");
            //    return (GridControl)gc;
            //    //if (tv != null)
            //    //{
            //    //    return (TableView)tv;
            //    //}
            //}

            return null;
        }

        private TableView FindTableView()
        {
            //if (_currentControl != null)
            //{
            //    var tv = _currentControl.FindName("tableView");

            //    if (tv != null)
            //    {
            //        return (TableView)tv;
            //    }
            //}

            return null;
        }

        protected void SaveDataToCSV(TableView tableView)
        {
            var dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.FileName = string.Format("{0} DataGrid", tableView.Tag);
            dlg.DefaultExt = ".xls";
            dlg.Filter = "CSV Documents (*.csv)|*.csv";

            Nullable<bool> result = dlg.ShowDialog();

            if (result == true)
            {
                tableView.ExportToCsv(dlg.FileName);
            }
        }

        protected void SaveDataToXLS(TableView tableView)
        {
            var dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.FileName = string.Format("{0} DataGrid", tableView.Tag);
            dlg.DefaultExt = ".xls";
            dlg.Filter = "XLS Documents (*.xls)|*.xls";

            Nullable<bool> result = dlg.ShowDialog();

            if (result == true)
            {
                tableView.ExportToXls(dlg.FileName);
            }
        }

        protected void SaveDataToXLSX(TableView tableView)
        {
            var dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.FileName = string.Format("{0} DataGrid", tableView.Tag);
            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "XLSX Documents (*.xlsx)|*.xlsx";

            Nullable<bool> result = dlg.ShowDialog();

            if (result == true)
            {
                tableView.ExportToXlsx(dlg.FileName);
            }
        }

        private void SetRowDetailType(TableView tableView)
        {
            //switch (Common.RowDetailMode)
            //{
            //    case "Tooltip":
            //        //case "tooltip":
            //        // This uses the resources from the usercontrol/window.
            //        //var  tmplt =  (DataTemplate)vtrootUserControl.Resources["rowTooltipDetailTemplate"];
            //        var tmplt = (DataTemplate)Application.Current.Resources["rowTooltipDetailTemplate"];
            //        tableView.DataRowTemplate = tmplt;
            //        break;

            //    case "SelectedRow":
            //        //case "selected":
            //        //tableView.DataRowTemplate = (DataTemplate)Resources["rowSelectedDetailTemplate"];
            //        // This uses the Application scope resources.
            //        var tmplt1 = (DataTemplate)Application.Current.Resources["rowSelectedDetailTemplate"];
            //        tableView.DataRowTemplate = tmplt1;
            //        break;

            //    case "DetailedRows":
            //        //case "details":
            //        // This uses FindResource.
            //        var tmplt2 = (DataTemplate)_currentControl.FindResource("rowDetailTemplate");
            //        tableView.DataRowTemplate = tmplt2;
            //        break;

            //    case "None":
            //        //case "none":
            //        tableView.ClearValue(DevExpress.Xpf.Grid.TableView.DataRowTemplateProperty);
            //        break;

            //    default:
            //        break;
            //}
        }

        private void SetTitle(object sender, EventArgs e)
        {
            //wucDXBase uc = sender as wucDXBase;

            //if (uc != null && !string.IsNullOrEmpty(uc.Title))
            //{
            //    this.Title = string.Format("VNCCodeCommandConsole: {0}", uc.Title);
            //}
            //else
            //{
            //    this.Title = "VNCCodeCommandConsole";
            //}
        }

        private void SetTitle2(object sender, EventArgs e)
        {
            //    wucDXBase uc = sender as wucDXBase;

            //    if (uc != null && !string.IsNullOrEmpty(uc.Title))
            //    {
            //        this.Title = string.Format("VNCCodeCommandConsole: {0}", uc.Title);
            //    }
            //    else
            //    {
            //        this.Title = "VNCCodeCommandConsole";
            //    }
        }

        private void ShowUserControl(UserControl control)
        {
            ////UnhookTitleEvent(_currentControl);
            //splashScreenGrid.Children.Clear();

            //if (control != null)
            //{
            //    splashScreenGrid.Children.Add(control);
            //    _currentControl = control;
            //    //_currentControl = (wucDX_Base)control;

            //    //if (control.MinWidth > 0)
            //    //{
            //    //    this.Width = control.MinWidth + Common.WINDOW_HOSTING_USER_CONTROL_WIDTH_PAD;
            //    //}

            //    //if (control.MinHeight > 0)
            //    //{
            //    //    this.Height = control.MinHeight + Common.WINDOW_HOSTING_USER_CONTROL_HEIGHT_PAD;
            //    //}

            //    TableView tableView = FindTableView();

            //    if (tableView != null)
            //    {
            //        SetRowDetailType(tableView);
            //    }
            //}

            ////HookTitleEvent(_currentControl);
        }

        //private void UnhookTitleEvent(wucDXBase control)
        //{
        //    if (control != null)
        //    {
        //        control.TitleChanged -= SetTitle;
        //    }
        //}

        //private void UpdateAvailableViews()
        //{
        //    //SplashScreenItems ssi = new SplashScreenItems();

        //    //if (Common.UserMode.Basic)
        //    //{
        //    //    lc_Root.Background = Brushes.Transparent;

        //    //    cbe_SplashScreens.ItemsSource = ssi.Items.Where(um => um.UserMode.Basic); ;
        //    //}
        //    //else if (Common.UserMode.Advanced)
        //    //{
        //    //    //lc_Root.Background = new SolidColorBrush(Color.FromArgb(0x10, 0x00, 0xFF, 0xC7));
        //    //    lc_Root.Background = Brushes.Transparent;

        //    //    cbe_SplashScreens.ItemsSource = ssi.Items.Where(um => um.UserMode.Advanced);
        //    //}
        //    //else if (Common.UserMode.Administrator)
        //    //{
        //    //    lc_Root.Background = new SolidColorBrush(Color.FromArgb(0x7E, 0xFF, 0x01, 0x01));

        //    //    cbe_SplashScreens.ItemsSource = ssi.Items.Where(um => um.UserMode.Administrator);
        //    //}
        //    //else if (Common.UserMode.Beta)
        //    //{
        //    //    cbe_SplashScreens.ItemsSource = ssi.Items.Where(um => um.UserMode.Beta);
        //    //}
        //}

        #endregion

        private void OnTakeSnapshot_Click(object sender, ItemClickEventArgs e)
        {
            //string snapShotType = (string)e.Item.Tag;

            //switch (snapShotType)
            //{
            //    case "Daily":
            //        SnapShot.TakeDaily();
            //        break;

            //    case "IntraDay":
            //        SnapShot.TakeIntraDay();
            //        break;

            //    case "Weekly":
            //        SnapShot.TakeWeekly();
            //        break;

            //    default:
            //        break;
            //}

            //MessageBox.Show(string.Format("Take{0}Snapshot", snapShotType));
        }

        private void bei_VisualTheme_EditValueChanged(object sender, RoutedEventArgs e)
        {
            //var bei = (BarEditItem)e.Source;

            //ApplicationThemeHelper.ApplicationThemeName = ((Theme)bei.EditValue).Name;

            //// Save it back to the config file.

            //Data.Config.DefaultUITheme = ApplicationThemeHelper.ApplicationThemeName;
        }

        #endregion
    }
}
