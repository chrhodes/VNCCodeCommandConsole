using System;
using System.Windows;

using VNC;
using VNC.Core.Mvvm;

namespace CCC.CodeChecks.Presentation.Views
{
    public partial class DesignChecks : ViewBase
    {
        public DesignChecks()
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InitializeComponent();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        //public DesignChecks(ViewModels.IChecksViewModel viewModel)
        //{
        //    Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

        //    InitializeComponent();

        //    ViewModel = viewModel;

        //    Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        //}
    }
}
