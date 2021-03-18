using System;
using System.Windows;

using VNC;
using VNC.Core.Mvvm;

namespace CCC.CodeChecks.Presentation.Views
{
    public partial class QualityChecks : ViewBase
    {
        public QualityChecks()
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InitializeComponent();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        //public QualityChecks(ViewModels.IChecksViewModel viewModel)
        //{
        //    Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

        //    InitializeComponent();

        //    ViewModel = viewModel;

        //    Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        //}
    }
}
