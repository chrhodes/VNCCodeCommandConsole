using System;

using VNC;
using VNC.Core.Mvvm;

namespace VNCCodeCommandConsole.Presentation.Views
{
    public partial class CatDetail : ViewBase, ICatDetail, IInstanceCountV
    {

        public CatDetail()
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InstanceCountV++;
            InitializeComponent();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public CatDetail(ViewModels.ICatDetailViewModel viewModel)
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

    }
}
