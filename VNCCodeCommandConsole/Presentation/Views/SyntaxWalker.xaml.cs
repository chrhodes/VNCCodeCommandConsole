using System;
using System.Windows;

using VNC;
using VNC.Core.Mvvm;

using VNCCodeCommandConsole.Presentation.ViewModels;

namespace VNCCodeCommandConsole.Presentation.Views
{
    public partial class SyntaxWalker : ViewBase, ISyntaxWalker, IInstanceCountV
    {

        public SyntaxWalker()
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InstanceCountV++;
            InitializeComponent();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public SyntaxWalker(ViewModels.ISyntaxWalkerViewModel viewModel)
        {
            Int64 startTicks = Log.CONSTRUCTOR($"Enter ({viewModel.GetType()})", Common.LOG_CATEGORY);

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

        private void ceCollapse_EditValueChanged(object sender, DevExpress.Xpf.Editors.EditValueChangedEventArgs e)
        {
            ((SyntaxWalkerViewModel)ViewModel).HeaderIsCollapsed = (bool)e.NewValue;
            ceCollapse.Content = $"{((bool)ceCollapse.IsChecked ? "Collapsed" : "Collapse")} Headers";
        }
    }
}
