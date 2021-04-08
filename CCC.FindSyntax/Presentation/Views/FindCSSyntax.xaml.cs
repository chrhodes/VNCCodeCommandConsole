using System;
using System.Windows;

using CCC.FindSyntax.Presentation.ViewModels;

using VNC;
using VNC.Core.Mvvm;

namespace CCC.FindSyntax.Presentation.Views
{
    public partial class FindCSSyntax : ViewBase, IInstanceCountV
    {
        public FindCSSyntax()
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            _instanceCountV++;
            InitializeComponent();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public FindCSSyntax(FindCSSyntaxViewModel viewModel)
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            _instanceCountV++;
            InitializeComponent();

            ViewModel = viewModel;

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void ceCollapse_EditValueChanged(object sender, DevExpress.Xpf.Editors.EditValueChangedEventArgs e)
        {
            ((FindCSSyntaxViewModel)ViewModel).HeaderIsCollapsed = (bool)e.NewValue;
            ceCollapse.Content = $"{((bool)ceCollapse.IsChecked ? "Collapsed" : "Collapse")} Headers";
        }

        #region IInstanceCount

        private static int _instanceCountV;

        public int InstanceCountV
        {
            get => _instanceCountV;
            set => _instanceCountV = value;
        }

        #endregion
    }
}
