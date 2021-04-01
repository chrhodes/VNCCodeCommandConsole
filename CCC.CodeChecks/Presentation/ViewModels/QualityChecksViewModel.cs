using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

using Prism.Commands;
using Prism.Events;

using VNC;
using VNC.Core.Events;
using VNC.Core.Mvvm;
using VNC.Core.Services;

using VNCCodeCommandConsole.Core.Events;

namespace CCC.CodeChecks.Presentation.ViewModels
{
    public class QualityChecksViewModel : EventViewModelBase, IInstanceCountVM
    {

        #region Constructors, Initialization, and Load

        public QualityChecksViewModel(
            IEventAggregator eventAggregator,
            IMessageDialogService messageDialogService) : base(eventAggregator, messageDialogService)
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            // TODO(crhodes)
            // Save constructor parameters here

            InitializeViewModel();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void InitializeViewModel()
        {
            Int64 startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            InstanceCountVM++;

            // TODO(crhodes)
            //

            CallCodeCheckCommand = new DelegateCommand<string>(
                CallCodeCheck, CallCodeCheckCanExecute);

            Message = "QualityChecksViewModel says hello";

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Enums


        #endregion

        #region Structures


        #endregion

        #region Fields and Properties

        public DelegateCommand<string> CallCodeCheckCommand { get; private set; }

        private string _language = "CS";
        private string _message;

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

        
        public string Language
        {
            get => _language;
            set
            {
                if (_language == value)
                    return;
                _language = value;
                OnPropertyChanged();
            }
        }
        

        #endregion

        #region Event Handlers


        #endregion

        #region Public Methods


        #endregion

        #region Protected Methods


        #endregion

        #region Private Methods

        private bool CallCodeCheckCanExecute(string codeCheckMethod)
        {
            return true;
        }

        void CallCodeCheck(string codeCheckMethod)
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            string targetName = codeCheckMethod;

            string language = Language;

            ////Boolean includeTrivia = ceStructuresIncludeTrivia.IsChecked.Value;
            ////Boolean statementsOnly = ceStructuresStatementsOnly.IsChecked.Value;

            StringBuilder sb = new StringBuilder();

            var sourceCode = "";

            //using (var sr = new StreamReader(CodeExplorerContext.teSourceFile.Text))
            //{
            //    sourceCode = sr.ReadToEnd();
            //}

            string metricClass = $"VNC.CodeAnalysis.QualityMetrics.{language}.{targetName},VNC.CodeAnalysis";

            // NOTE(crhodes)
            // I think we can just pass the metricClass in the event

            EventAggregator.GetEvent<InvokeCodeCheckEvent>().Publish(metricClass);
            Message = metricClass;

            //Type metricType = Type.GetType(metricClass);
            //MethodInfo metricMethod = metricType.GetMethod("Check");
            //object[] parametersArray = new object[] { sourceCode };

            //sb = (StringBuilder)metricMethod.Invoke(null, parametersArray);

            //CodeExplorer.teSourceCode.Text = sb.ToString();

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region IInstanceCount

        private static int _instanceCountVM;

        public int InstanceCountVM
        {
            get => _instanceCountVM;
            set => _instanceCountVM = value;
        }

        #endregion IInstanceCount
    }
}
