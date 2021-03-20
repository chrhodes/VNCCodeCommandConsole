using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml.Linq;

using Microsoft.Win32;

namespace VNCCodeCommandConsole.User_Interface.User_Controls
{
    public partial class wucSourceBranch_Picker : UserControl
    {

        #region Constructors and Load

        public wucSourceBranch_Picker()
        {
            InitializeComponent();
        }

        #endregion

        #region Properties and Fields

        public XElement xElement;

        // TODO:
        //  Add properties, backing fields, and dependency properties to match the attributes from each XML element
        //  Convention is _<Attribute Name>, <AttributeName>, <AttributeName>DP
        //  where <AttributeName> is the name of the attribute for the <Element>

        private string _Name;

        public string Name
        {
            get { return _Name; }
            set
            {
                _Name = value;
            }
        }

        public string NameDP
        {
            // IMPORTANT: To maintain parity between setting a property in XAML and procedural code, do not touch the getter and setter inside this dependency property!
            get
            {
                return (string)GetValue(NameDPProperty);
            }
            set
            {
                SetValue(NameDPProperty, value);
            }
        }

        // TODO:
        //  Update the parameters to .Register(...) as needed

        public static readonly DependencyProperty NameDPProperty = DependencyProperty.Register(
            "NameDP", typeof(string), typeof(wucSourceBranch_Picker), new PropertyMetadata(null, new PropertyChangedCallback(OnNameDPChanged)));

        private string _Repository;

        public string Repository
        {
            get { return _Repository; }
            set
            {
                _Repository = value;
            }
        }

        public string RepositoryDP
        {
            // IMPORTANT: To maintain parity between setting a property in XAML and procedural code, do not touch the getter and setter inside this dependency property!
            get
            {
                return (string)GetValue(RepositoryDPProperty);
            }
            set
            {
                SetValue(RepositoryDPProperty, value);
            }
        }

        // TODO:
        //  Update the parameters to .Register(...) as needed

        public static readonly DependencyProperty RepositoryDPProperty = DependencyProperty.Register(
            "RepositoryDP", typeof(string), typeof(wucSourceBranch_Picker), new PropertyMetadata(null, new PropertyChangedCallback(OnRepositoryDPChanged)));

        public string SourcePath
        {
            get { return _SourcePath; }
            set
            {
                _SourcePath = value;
            }
        }

        private string _SourcePath;
        public string SourcePathDP
        {
            // IMPORTANT: To maintain parity between setting a property in XAML and procedural code, do not touch the getter and setter inside this dependency property!
            get
            {
                return (string)GetValue(SourcePathDPProperty);
            }
            set
            {
                SetValue(SourcePathDPProperty, value);
            }
        }
        // TODO:
        //  Update the parameters to .Register(...) as needed

        public static readonly DependencyProperty SourcePathDPProperty = DependencyProperty.Register(
            "SourcePathDP", typeof(string), typeof(wucSourceBranch_Picker), new PropertyMetadata(null, new PropertyChangedCallback(OnSourcePathDPChanged)));

        #endregion

        #region Delegates and Events

        public delegate void ControlEvent();
        public event ControlEvent ControlChanged;

        #endregion

        #region Event Handlers
        private void ComboBoxEdit_EditValueChanged(object sender, DevExpress.Xpf.Editors.EditValueChangedEventArgs e)
        {
            var v = sender;
            var item = e.NewValue;
            //var item = ((System.Windows.Controls.ComboBox)v).SelectedItem;
            System.Xml.XmlElement xmlElement = (System.Xml.XmlElement)item;
            xElement = XElement.Parse(xmlElement.OuterXml);

            if (null == item)
            {
                // May have just opened new file and no item has been selected.
                return;
            }

            _Name = xmlElement.Attributes["Name"].Value;
            NameDP = xmlElement.Attributes["Name"].Value;

            _Repository = xmlElement.Attributes["Repository"].Value;
            RepositoryDP = xmlElement.Attributes["Repository"].Value;

            _SourcePath = xmlElement.Attributes["SourcePath"].Value;
            SourcePathDP = xmlElement.Attributes["SourcePath"].Value;

            ControlEvent fireEvent = Interlocked.CompareExchange(ref ControlChanged, null, null);

            if (null != fireEvent)
            {
                fireEvent();
            }
        }

        protected virtual void OnNameDPChanged(string oldValue, string newValue)
        {
            // TODO: Add your property changed side-effects. Descendants can override as well.
        }

        private static void OnNameDPChanged(DependencyObject o, DependencyPropertyChangedEventArgs e)
        {
            wucSourceBranch_Picker picker = o as wucSourceBranch_Picker;
            if (picker != null)
                picker.OnNameDPChanged((string)e.OldValue, (string)e.NewValue);
        }

        protected virtual void OnRepositoryDPChanged(string oldValue, string newValue)
        {
            // TODO: Add your property changed side-effects. Descendants can override as well.
        }

        private static void OnRepositoryDPChanged(DependencyObject o, DependencyPropertyChangedEventArgs e)
        {
            wucSourceBranch_Picker picker = o as wucSourceBranch_Picker;
            if (picker != null)
                picker.OnRepositoryDPChanged((string)e.OldValue, (string)e.NewValue);
        }
        protected virtual void OnSourcePathDPChanged(string oldValue, string newValue)
        {
            // TODO: Add your property changed side-effects. Descendants can override as well.
        }

        private static void OnSourcePathDPChanged(DependencyObject o, DependencyPropertyChangedEventArgs e)
        {
            wucSourceBranch_Picker picker = o as wucSourceBranch_Picker;
            if (picker != null)
                picker.OnSourcePathDPChanged((string)e.OldValue, (string)e.NewValue);
        }

        private void Reload_Click(object sender, RoutedEventArgs e)
        {
            LoadDataFromFile();
        }

        #endregion

        #region Main Methods

        public void PopulateControlFromFile(string fileNameAndPath)
        {
            comboBox.Source = new Uri(fileNameAndPath);
        }

        #endregion

        #region Private Methods

        private void LoadDataFromFile()
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*";
            openFileDialog1.FileName = "";

            if (true == openFileDialog1.ShowDialog())
            {
                string fileName = openFileDialog1.FileName;

                PopulateControlFromFile(fileName);
            }
        }

        #endregion
    }
}
