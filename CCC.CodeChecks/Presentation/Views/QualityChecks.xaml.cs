using System;
using System.Windows;

using VNC;
using VNC.Core.Mvvm;

namespace CCC.CodeChecks.Presentation.Views
{
    public partial class QualityChecks : ViewBase, IInstanceCountV
    {
        public QualityChecks()
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            _instanceCountV++;
            InitializeComponent();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public QualityChecks(ViewModels.QualityChecksViewModel viewModel)
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
    }
}
