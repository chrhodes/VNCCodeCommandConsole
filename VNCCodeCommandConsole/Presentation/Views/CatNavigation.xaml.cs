using System;

using VNC;
using VNC.Core.Mvvm;

namespace VNCCodeCommandConsole.Presentation.Views
{
    public partial class CatNavigation : ViewBase, ICatNavigation, IInstanceCountV
    {
        public CatNavigation()
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            InstanceCountV++;
            InitializeComponent();

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
