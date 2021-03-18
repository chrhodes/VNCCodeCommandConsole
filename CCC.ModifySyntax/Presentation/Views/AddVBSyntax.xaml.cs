using System;
using System.Windows;

using VNC;
using VNC.Core.Mvvm;

namespace CCC.ModifySyntax.Presentation.Views
{
    public partial class AddVBSyntax : ViewBase
    {
        public AddVBSyntax()
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InitializeComponent();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        //public AddVBSyntax(ViewModels.IFindSyntaxViewModel viewModel)
        //{
        //    Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

        //    InitializeComponent();

        //    ViewModel = viewModel;

        //    Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        //}
    }
}
