using System;
using System.Windows;

using VNC;
using VNC.Core.Mvvm;

using CCC.CodeChecks.VB.Presentation.Views;

namespace CCC.CodeChecks.Presentation.Views
{
    public partial class DesignChecks : ViewBase, IDesignChecks, IInstanceCountV
    {
        public DesignChecks()
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            _instanceCountV++;
            InitializeComponent();
            //((IViewModel)DataContext).View = this;

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        //public DesignChecks(ViewModels.DesignChecksViewModel viewModel)
        public DesignChecks(ViewModels.IDesignChecksViewModel viewModel)
        {
            Int64 startTicks = Log.CONSTRUCTOR($"Enter viewModel({viewModel.GetType()}", Common.LOG_CATEGORY);

            _instanceCountV++;
            InitializeComponent();

            ViewModel = viewModel;
            viewModel.View = this;

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
