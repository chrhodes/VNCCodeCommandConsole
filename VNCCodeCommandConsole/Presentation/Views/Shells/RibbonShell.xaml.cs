﻿using System;
using System.Windows;

using VNC;
using VNC.Core.Mvvm;

using VNCCodeCommandConsole.Presentation.ViewModels;

namespace VNCCodeCommandConsole.Presentation.Views
{
    public partial class RibbonShell : Window, IInstanceCountV
    {
        public RibbonShellViewModel _viewModel;

        public RibbonShell(RibbonShellViewModel viewModel)
        {
            Int64 startTicks = Log.CONSTRUCTOR($"Enter ({viewModel.GetType()})", Common.LOG_CATEGORY);

            InstanceCountV++;
            InitializeComponent();

            _viewModel = viewModel;
            DataContext = _viewModel;

            Log.CONSTRUCTOR(String.Format("Exit"), Common.LOG_CATEGORY, startTicks);
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
