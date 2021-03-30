using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using CCC.FindSyntax.Presentation.ViewModels;

using Prism.Commands;

using VNC;
using VNC.Core.Mvvm;

namespace CCC.FindSyntax.Presentation.Views
{
    /// <summary>
    /// Interaction logic for RegExStateWalkerBase.xaml
    /// </summary>
    public partial class RegExSyntaxWalker : VNC.Core.Mvvm.ViewBase, IInstanceCountV, INotifyPropertyChanged
    {
        public RegExSyntaxWalker()
        {
            InitializeComponent();

            SyntaxWalkerCommand = new DelegateCommand(
                WalkerExecute, WalkerCanExecute);

            lgRegularExpression.DataContext = this;
            //DataContext = this;
        }

        public event PropertyChangedEventHandler PropertyChanged;


        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        string _message = "RESW-V-InitialMessage";


        public string Message
        {
            get => _message;
            set
            {
                if (_message == value)
                    return;
                _message = value;
                OnPropertyChanged();
            }
        }

        public DelegateCommand SyntaxWalkerCommand { get; set; }

        public void WalkerExecute()
        {
            Int64 startTicks = Log.EVENT("Enter", Common.LOG_CATEGORY);

            //Helper.ProcessOperation(DisplayImportsStatementWalkerVB, CodeExplorer, CodeExplorerContext, CodeExplorer.configurationOptions);

            Message = $"RESW-V-WE {DateTime.Now.ToLongTimeString()} RegEx:({RegEx})";

            //EventAggregator.GetEvent<InvokeVBSyntaxWalkerEvent>().Publish(SearchTreeCommand);

            //EventAggregator.GetEvent<SyntaxWalkerResultEvent>().Publish("Syntax Walker Result from ImportsStatementWalker");

            // If using events to tell something else to act

            //    EventAggregator.GetEvent<ImportsStatementWalkerEvent>().Publish(
            //        new ImportsStatementWalkerEventArgs()
            //        {
            //             Organization = _collectionMainViewModel.SelectedCollection.Organization,
            //             Project = _contextMainViewModel.Context.SelectedProject,
            //             Team = _contextMainViewModel.Context.SelectedTeam
            //        });

            // Put this in places that listen for event and create method ImportsStatementWalker
            // EventAggregator.GetEvent<ImportsStatementWalkerEvent>().Subscribe(ImportsStatementWalker);

            Log.EVENT("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public bool WalkerCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        public RegExSyntaxWalker(ViewModels.RegExSyntaxWalkerViewModel viewModel)
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            _instanceCountV++;
            InitializeComponent();

            ViewModel = viewModel;

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public static readonly DependencyProperty RegExProperty = DependencyProperty.Register(
            "RegEx", typeof(string), typeof(RegExSyntaxWalker),
            new PropertyMetadata("RESW-V.*", new PropertyChangedCallback(OnRegExChanged)));


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
            // IMPORTANT: To maintain parity between setting a property in XAML and procedural code, do not touch the getter and setter inside this dependency property!
            get => (bool)GetValue(UseRegExProperty);
            set => SetValue(UseRegExProperty, value);
        }

        public static readonly DependencyProperty ButtonContentProperty = DependencyProperty.Register(
            "ButtonContent", typeof(string), typeof(RegExSyntaxWalker),
            new PropertyMetadata("DEFAULT BUTTON", new PropertyChangedCallback(OnButtonContentChanged)));

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
            // IMPORTANT: To maintain parity between setting a property in XAML and procedural code, do not touch the getter and setter inside this dependency property!
            get => (string)GetValue(ButtonContentProperty);
            set => SetValue(ButtonContentProperty, value);
        }

        public static readonly DependencyProperty CommandParameterProperty = DependencyProperty.Register(
            "CommandParameter", typeof(string), typeof(RegExSyntaxWalker), 
            new PropertyMetadata("DEFAULTWALKER", new PropertyChangedCallback(OnCommandParameterChanged)));

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
            // IMPORTANT: To maintain parity between setting a property in XAML and procedural code, do not touch the getter and setter inside this dependency property!
            get => (string)GetValue(CommandParameterProperty);
            set => SetValue(CommandParameterProperty, value);
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
