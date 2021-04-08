using System;
using System.Windows;

using CCC.FindSyntax.Presentation.ViewModels;

using VNC;
using VNC.Core.Mvvm;

namespace CCC.FindSyntax.Presentation.Views
{
    public partial class FindVBSyntax : ViewBase, IInstanceCountV
    {
        public FindVBSyntax()
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            _instanceCountV++;
            InitializeComponent();
            var foo = cImportsStatement;
            var bar = cImportsStatement.DataContext;
            //cImportsStatement.TheButton.Command = bar.            
            //(cImportsStatement.DataContext) = "FooBar";

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public FindVBSyntax(FindVBSyntaxViewModel viewModel)
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            _instanceCountV++;
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

        private void ceCollapse_EditValueChanged(object sender, DevExpress.Xpf.Editors.EditValueChangedEventArgs e)
        {
            ((FindVBSyntaxViewModel)ViewModel).HeaderIsCollapsed = (bool)e.NewValue;
            ceCollapse.Content = $"{((bool)ceCollapse.IsChecked ? "Collapsed" : "Collapse")} Headers";
        }
    }
}
