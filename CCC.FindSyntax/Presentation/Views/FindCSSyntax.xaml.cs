using System;
using System.Windows;

using VNC;
using VNC.Core.Mvvm;

namespace CCC.FindSyntax.Presentation.Views
{
    public partial class FindCSSyntax : ViewBase
    {
        public FindCSSyntax()
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InitializeComponent();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        //public FindCSSyntax(ViewModels.IFindSyntaxViewModel viewModel)
        //{
        //    Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

        //    InitializeComponent();

        //    ViewModel = viewModel;

        //    Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        //}
    }
}
