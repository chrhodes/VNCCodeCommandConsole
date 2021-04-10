using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;

using VNC.Core.Mvvm;

namespace CCC.FindSyntax.Presentation.Views
{
    public partial class RegExStructSyntaxWalker : VNC.Core.Mvvm.ViewBase, IInstanceCountV, INotifyPropertyChanged
    {
        public RegExStructSyntaxWalker()
        {
            InitializeComponent();

            // Need to do this so UseRegEx and RegEx fire notification events.
            lgRegularExpression.DataContext = this;
            lgRegularExpressionFields.DataContext = this;
            ceShowFields.DataContext = this;
        }

        public event PropertyChangedEventHandler PropertyChanged;


        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #region Dependency Properties

        #region ControlHeader

        public static readonly DependencyProperty ControlHeaderProperty = DependencyProperty.Register(
            "ControlHeader", typeof(string), typeof(RegExStructSyntaxWalker),
            new PropertyMetadata("<HEADER>", new PropertyChangedCallback(OnControlHeaderChanged)));

        private static void OnControlHeaderChanged(DependencyObject o, DependencyPropertyChangedEventArgs e)
        {
            RegExStructSyntaxWalker RegExStructSyntaxWalker = o as RegExStructSyntaxWalker;
            if (RegExStructSyntaxWalker != null)
                RegExStructSyntaxWalker.OnControlHeaderChanged((string)e.OldValue, (string)e.NewValue);
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
            "RegExLabel", typeof(string), typeof(RegExStructSyntaxWalker),
            new PropertyMetadata("RegEx", new PropertyChangedCallback(OnRegExLabelChanged)));

        private static void OnRegExLabelChanged(DependencyObject o, DependencyPropertyChangedEventArgs e)
        {
            RegExStructSyntaxWalker RegExStructSyntaxWalker = o as RegExStructSyntaxWalker;
            if (RegExStructSyntaxWalker != null)
                RegExStructSyntaxWalker.OnRegExLabelChanged((string)e.OldValue, (string)e.NewValue);
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
            "RegEx", typeof(string), typeof(RegExStructSyntaxWalker),
            new PropertyMetadata("RESW-V.*", new PropertyChangedCallback(OnRegExChanged)));


        private static void OnRegExChanged(DependencyObject o, DependencyPropertyChangedEventArgs e)
        {
            RegExStructSyntaxWalker RegExStructSyntaxWalker = o as RegExStructSyntaxWalker;
            if (RegExStructSyntaxWalker != null)
                RegExStructSyntaxWalker.OnRegExChanged((string)e.OldValue, (string)e.NewValue);
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
            typeof(bool), typeof(RegExStructSyntaxWalker),
            new PropertyMetadata(false, new PropertyChangedCallback(OnUseRegExChanged)));

        private static void OnUseRegExChanged(DependencyObject o, DependencyPropertyChangedEventArgs e)
        {
            RegExStructSyntaxWalker RegExStructSyntaxWalker = o as RegExStructSyntaxWalker;
            if (RegExStructSyntaxWalker != null)
                RegExStructSyntaxWalker.OnUseRegExChanged((bool)e.OldValue, (bool)e.NewValue);
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
            "ButtonContent", typeof(string), typeof(RegExStructSyntaxWalker),
            new PropertyMetadata("<DEFAULT BUTTON>", new PropertyChangedCallback(OnButtonContentChanged)));

        private static void OnButtonContentChanged(DependencyObject o, DependencyPropertyChangedEventArgs e)
        {
            RegExStructSyntaxWalker RegExStructSyntaxWalker = o as RegExStructSyntaxWalker;
            if (RegExStructSyntaxWalker != null)
                RegExStructSyntaxWalker.OnButtonContentChanged((string)e.OldValue, (string)e.NewValue);
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
            "CommandParameter", typeof(string), typeof(RegExStructSyntaxWalker),
            new PropertyMetadata("<DEFAULTWALKER>", new PropertyChangedCallback(OnCommandParameterChanged)));

        private static void OnCommandParameterChanged(DependencyObject o, DependencyPropertyChangedEventArgs e)
        {
            RegExStructSyntaxWalker RegExStructSyntaxWalker = o as RegExStructSyntaxWalker;
            if (RegExStructSyntaxWalker != null)
                RegExStructSyntaxWalker.OnCommandParameterChanged((string)e.OldValue, (string)e.NewValue);
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
            "HeaderIsCollapsed", typeof(bool), typeof(RegExStructSyntaxWalker),
            new PropertyMetadata(true, new PropertyChangedCallback(OnHeaderIsCollapsedChanged)));

        private static void OnHeaderIsCollapsedChanged(DependencyObject o, DependencyPropertyChangedEventArgs e)
        {
            RegExStructSyntaxWalker regExSyntaxWalker = o as RegExStructSyntaxWalker;
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

        public static readonly DependencyProperty DisplayFieldsProperty = DependencyProperty.Register(
            "DisplayFields", typeof(bool), typeof(RegExStructSyntaxWalker), 
            new PropertyMetadata(false, new PropertyChangedCallback(OnDisplayFieldsChanged)));

        private static void OnDisplayFieldsChanged(DependencyObject o, DependencyPropertyChangedEventArgs e)
        {
            RegExStructSyntaxWalker regExStructSyntaxWalker = o as RegExStructSyntaxWalker;
            if (regExStructSyntaxWalker != null)
                regExStructSyntaxWalker.OnDisplayFieldsChanged((bool)e.OldValue, (bool)e.NewValue);
        }

        protected virtual void OnDisplayFieldsChanged(bool oldValue, bool newValue)
        {
            // TODO: Add your property changed side-effects. Descendants can override as well.
        }

        public bool DisplayFields
        {
            // IMPORTANT: To maintain parity between setting a property in XAML and procedural code, do not touch the getter and setter inside this dependency property!
            get => (bool)GetValue(DisplayFieldsProperty);
            set => SetValue(DisplayFieldsProperty, value);
        }

        #region RegExFields

        public static readonly DependencyProperty RegExFieldsProperty = DependencyProperty.Register(
            "RegExFields", typeof(string), typeof(RegExStructSyntaxWalker),
            new PropertyMetadata(".*", new PropertyChangedCallback(OnRegExFieldsChanged)));


        private static void OnRegExFieldsChanged(DependencyObject o, DependencyPropertyChangedEventArgs e)
        {
            RegExStructSyntaxWalker RegExStructSyntaxWalker = o as RegExStructSyntaxWalker;
            if (RegExStructSyntaxWalker != null)
                RegExStructSyntaxWalker.OnRegExFieldsChanged((string)e.OldValue, (string)e.NewValue);
        }

        protected virtual void OnRegExFieldsChanged(string oldValue, string newValue)
        {
            // TODO: Add your property changed side-effects. Descendants can override as well.
        }

        public string RegExFields
        {
            // IMPORTANT: To maintain parity between setting a property in XAML and procedural code, do not touch the getter and setter inside this dependency property!
            get => (string)GetValue(RegExFieldsProperty);
            set => SetValue(RegExFieldsProperty, value);
        }

        #endregion

        #region UseRegExFields

        public static readonly DependencyProperty UseRegExFieldsProperty = DependencyProperty.Register(
            "UseRegExFields", typeof(bool), typeof(RegExStructSyntaxWalker),
            new PropertyMetadata(false, new PropertyChangedCallback(OnUseRegExFieldsChanged)));

        private static void OnUseRegExFieldsChanged(DependencyObject o, DependencyPropertyChangedEventArgs e)
        {
            RegExStructSyntaxWalker RegExStructSyntaxWalker = o as RegExStructSyntaxWalker;
            if (RegExStructSyntaxWalker != null)
                RegExStructSyntaxWalker.OnUseRegExFieldsChanged((bool)e.OldValue, (bool)e.NewValue);
        }

        protected virtual void OnUseRegExFieldsChanged(bool oldValue, bool newValue)
        {
            // TODO: Add your property changed side-effects. Descendants can override as well.
        }

        public bool UseRegExFields
        {
            // IMPORTANT: To maintain parity between setting a property in XAML and procedural code,
            // do not touch the getter and setter inside this dependency property!
            get => (bool)GetValue(UseRegExFieldsProperty);
            set => SetValue(UseRegExFieldsProperty, value);
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
