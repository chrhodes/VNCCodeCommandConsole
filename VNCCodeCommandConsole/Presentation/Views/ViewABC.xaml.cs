using System;
using System.Windows;

using VNC;
using VNC.Core.Mvvm;

namespace VNCCodeCommandConsole.Presentation.Views
{
    public partial class ViewABC : ViewBase, IViewABC, IInstanceCountV
    {

        public ViewABC(ViewModels.IViewABCViewModel viewModel)
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

    }
}
