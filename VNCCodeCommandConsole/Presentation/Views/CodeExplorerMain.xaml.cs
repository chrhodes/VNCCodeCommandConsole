using System;
using System.Windows;

using VNC;
using VNC.Core.Mvvm;

using VNCCodeCommandConsole.Presentation.ViewModels;

namespace VNCCodeCommandConsole.Presentation.Views
{
    public partial class CodeExplorerMain : ViewBase, IMain, IInstanceCountV
    {
        public ICodeExplorerMainViewModel _viewModel;

        public CodeExplorerMain(ICodeExplorerMainViewModel viewModel)
        {
            Int64 startTicks = Log.CONSTRUCTOR($"Enter viewModel({viewModel.GetType()})", Common.LOG_CATEGORY);

            InstanceCountV++;
            InitializeComponent();

            _viewModel = viewModel;
            DataContext = _viewModel;

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
