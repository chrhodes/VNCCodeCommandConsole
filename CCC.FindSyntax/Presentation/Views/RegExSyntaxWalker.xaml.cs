﻿using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;

using VNC.Core.Mvvm;

namespace CCC.FindSyntax.Presentation.Views
{
    public partial class RegExSyntaxWalker : VNC.Core.Mvvm.ViewBase, IInstanceCountV, INotifyPropertyChanged
    {
        public RegExSyntaxWalker()
        {
            InitializeComponent();

            // Need to do this so UseRegEx and RegEx fire notification events.
            lgRegularExpression.DataContext = this;

            // Cannot do this as it breaks the button
            //lgHeader.DataContext = this;

        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #region Dependency Properties

        #region ControlHeader

        public static readonly DependencyProperty ControlHeaderProperty = DependencyProperty.Register(
            "ControlHeader", typeof(string), typeof(RegExSyntaxWalker),
            new PropertyMetadata("<HEADER>", new PropertyChangedCallback(OnControlHeaderChanged)));

        private static void OnControlHeaderChanged(DependencyObject o, DependencyPropertyChangedEventArgs e)
        {
            RegExSyntaxWalker regExSyntaxWalker = o as RegExSyntaxWalker;
            if (regExSyntaxWalker != null)
                regExSyntaxWalker.OnControlHeaderChanged((string)e.OldValue, (string)e.NewValue);
        }

        protected virtual void OnControlHeaderChanged(string oldValue, string newValue)
        {
            lgHeader.Header = newValue;
            // TODO: Add your property changed side-effects. Descendants can override as well.
        }

        public string ControlHeader
        {
            // IMPORTANT: To maintain parity between setting a property in XAML and procedural code, do not touch the getter and setter inside this dependency property!
            get => (string)GetValue(ControlHeaderProperty);
            set => SetValue(ControlHeaderProperty, value);
        }

        #endregion

        #region RegExLabel

        public static readonly DependencyProperty RegExLabelProperty = DependencyProperty.Register(
            "RegExLabel", typeof(string), typeof(RegExSyntaxWalker),
            new PropertyMetadata("RegEx", new PropertyChangedCallback(OnRegExLabelChanged)));

        private static void OnRegExLabelChanged(DependencyObject o, DependencyPropertyChangedEventArgs e)
        {
            RegExSyntaxWalker regExSyntaxWalker = o as RegExSyntaxWalker;
            if (regExSyntaxWalker != null)
                regExSyntaxWalker.OnRegExLabelChanged((string)e.OldValue, (string)e.NewValue);
        }

        protected virtual void OnRegExLabelChanged(string oldValue, string newValue)
        {
            lblRegEx.Label = newValue;
            // TODO: Add your property changed side-effects. Descendants can override as well.
        }

        public string RegExLabel
        {
            // IMPORTANT: To maintain parity between setting a property in XAML and procedural code,
            // do not touch the getter and setter inside this dependency property!
            get => (string)GetValue(RegExLabelProperty);
            set => SetValue(RegExLabelProperty, value);
        }

        #endregion

        #region RegEx

        public static readonly DependencyProperty RegExProperty = DependencyProperty.Register(
            "RegEx", typeof(string), typeof(RegExSyntaxWalker),
            new PropertyMetadata(".*", new PropertyChangedCallback(OnRegExChanged)));

        private static void OnRegExChanged(DependencyObject o, DependencyPropertyChangedEventArgs e)
        {
            RegExSyntaxWalker regExSyntaxWalker = o as RegExSyntaxWalker;
            if (regExSyntaxWalker != null)
                regExSyntaxWalker.OnRegExChanged((string)e.OldValue, (string)e.NewValue);
        }

        protected virtual void OnRegExChanged(string oldValue, string newValue)
        {
            // TODO: Add your property changed side-effects. Descendants can override as well.
        }

        public string RegEx
        {
            // IMPORTANT: To maintain parity between setting a property in XAML and procedural code, do not touch the getter and setter inside this dependency property!
            get => (string)GetValue(RegExProperty);
            set => SetValue(RegExProperty, value);
        }

        #endregion

        #region UseRegEx

        public static readonly DependencyProperty UseRegExProperty = DependencyProperty.Register("UseRegEx",
            typeof(bool), typeof(RegExSyntaxWalker),
            new PropertyMetadata(false, new PropertyChangedCallback(OnUseRegExChanged)));

        private static void OnUseRegExChanged(DependencyObject o, DependencyPropertyChangedEventArgs e)
        {
            RegExSyntaxWalker regExSyntaxWalker = o as RegExSyntaxWalker;
            if (regExSyntaxWalker != null)
                regExSyntaxWalker.OnUseRegExChanged((bool)e.OldValue, (bool)e.NewValue);
        }

        protected virtual void OnUseRegExChanged(bool oldValue, bool newValue)
        {
            // TODO: Add your property changed side-effects. Descendants can override as well.
        }

        public bool UseRegEx
        {
            // IMPORTANT: To maintain parity between setting a property in XAML and procedural code,
            // do not touch the getter and setter inside this dependency property!
            get => (bool)GetValue(UseRegExProperty);
            set => SetValue(UseRegExProperty, value);
        }

        #endregion

        #region ButtonContent

        public static readonly DependencyProperty ButtonContentProperty = DependencyProperty.Register(
            "ButtonContent", typeof(string), typeof(RegExSyntaxWalker),
            new PropertyMetadata("<DEFAULT BUTTON>", new PropertyChangedCallback(OnButtonContentChanged)));

        private static void OnButtonContentChanged(DependencyObject o, DependencyPropertyChangedEventArgs e)
        {
            RegExSyntaxWalker regExSyntaxWalker = o as RegExSyntaxWalker;
            if (regExSyntaxWalker != null)
                regExSyntaxWalker.OnButtonContentChanged((string)e.OldValue, (string)e.NewValue);
        }

        protected virtual void OnButtonContentChanged(string oldValue, string newValue)
        {
            btnButton.Content = newValue;
            // TODO: Add your property changed side-effects. Descendants can override as well.
        }

        public string ButtonContent
        {
            // IMPORTANT: To maintain parity between setting a property in XAML and procedural code,
            // do not touch the getter and setter inside this dependency property!
            get => (string)GetValue(ButtonContentProperty);
            set => SetValue(ButtonContentProperty, value);
        }

        #endregion

        #region CommandParameter

        public static readonly DependencyProperty CommandParameterProperty = DependencyProperty.Register(
            "CommandParameter", typeof(string), typeof(RegExSyntaxWalker),
            new PropertyMetadata("<DEFAULTWALKER>", new PropertyChangedCallback(OnCommandParameterChanged)));

        private static void OnCommandParameterChanged(DependencyObject o, DependencyPropertyChangedEventArgs e)
        {
            RegExSyntaxWalker regExSyntaxWalker = o as RegExSyntaxWalker;
            if (regExSyntaxWalker != null)
                regExSyntaxWalker.OnCommandParameterChanged((string)e.OldValue, (string)e.NewValue);
        }

        protected virtual void OnCommandParameterChanged(string oldValue, string newValue)
        {
            btnButton.CommandParameter = newValue;
            // TODO: Add your property changed side-effects. Descendants can override as well.
        }

        public string CommandParameter
        {
            // IMPORTANT: To maintain parity between setting a property in XAML and procedural code,
            // do not touch the getter and setter inside this dependency property!
            get => (string)GetValue(CommandParameterProperty);
            set => SetValue(CommandParameterProperty, value);
        }

        #endregion

        #region HeaderIsCollapsed

        public static readonly DependencyProperty HeaderIsCollapsedProperty = DependencyProperty.Register(
            "HeaderIsCollapsed", typeof(bool), typeof(RegExSyntaxWalker), 
            new PropertyMetadata(true, new PropertyChangedCallback(OnHeaderIsCollapsedChanged)));

        private static void OnHeaderIsCollapsedChanged(DependencyObject o, DependencyPropertyChangedEventArgs e)
        {
            RegExSyntaxWalker regExSyntaxWalker = o as RegExSyntaxWalker;
            if (regExSyntaxWalker != null)
                regExSyntaxWalker.OnHeaderIsCollapsedChanged((bool)e.OldValue, (bool)e.NewValue);
        }

        protected virtual void OnHeaderIsCollapsedChanged(bool oldValue, bool newValue)
        {
            lgHeader.IsCollapsed = newValue;
            // TODO: Add your property changed side-effects. Descendants can override as well.
        }

        public bool HeaderIsCollapsed
        {
            // IMPORTANT: To maintain parity between setting a property in XAML and procedural code,
            // do not touch the getter and setter inside this dependency property!
            get => (bool)GetValue(HeaderIsCollapsedProperty);
            set => SetValue(HeaderIsCollapsedProperty, value);
        }

        #endregion

        #endregion

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
