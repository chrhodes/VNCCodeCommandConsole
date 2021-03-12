﻿using System;
using System.Windows;

using VNC;

using VNCCodeCommandConsole.Presentation.ViewModels;

namespace VNCCodeCommandConsole.Presentation.Views
{
    public partial class Shell : Window
    {
        public ShellViewModel _viewModel;

        public Shell(ShellViewModel viewModel)
        {
            Int64 startTicks = Log.CONSTRUCTOR($"Enter ({viewModel.GetType()})", Common.LOG_APPNAME);

            InitializeComponent();

            _viewModel = viewModel;
            DataContext = _viewModel;

            Log.CONSTRUCTOR(String.Format("Exit"), Common.LOG_APPNAME, startTicks);
        }
    }
}
