using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Input;

using Prism.Commands;
using Prism.Events;

using VNC;
using VNC.Core.Events;
using VNC.Core.Mvvm;
using VNC.Core.Services;

using VNCCodeCommandConsole.Domain;
using VNCCodeCommandConsole.DomainServices;
using VNCCodeCommandConsole.Presentation.ModelWrappers;

namespace VNCCodeCommandConsole.Presentation.ViewModels
{
    public class ToyDetailViewModel : DetailViewModelBase, IToyDetailViewModel, IInstanceCountVM
    {
        #region Contructors, Initialization, and Load

        public ToyDetailViewModel(
            IToyDataService ToyDataService,
            IEventAggregator eventAggregator,
            IMessageDialogService messageDialogService) : base(eventAggregator, messageDialogService)
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_APPNAME);

            _ToyDataService = ToyDataService;

            InitializeViewModel();

            Log.CONSTRUCTOR("Exit", Common.LOG_APPNAME, startTicks);
        }

        private void InitializeViewModel()
        {
            Int64 startTicks = Log.VIEWMODEL("Enter", Common.LOG_APPNAME);

            InstanceCountVM++;

            Title = "Toys";
            Toys = new ObservableCollection<ToyWrapper>();

            AddCommand = new DelegateCommand(AddExecute);
            RemoveCommand = new DelegateCommand(RemoveExecute, RemoveCanExecute);

            Log.VIEWMODEL("Exit", Common.LOG_APPNAME, startTicks);
        }

        #endregion

        #region Enums

        #endregion

        #region Structures

        #endregion

        #region Fields and Properties

        private IToyDataService _ToyDataService;

        public ICommand AddCommand { get; private set; }

        public ICommand RemoveCommand { get; private set; }

        public ObservableCollection<ToyWrapper> Toys { get; private set; }

        private ToyWrapper _selectedToy;

        public ToyWrapper SelectedToy
        {
            get { return _selectedToy; }
            set
            {
                _selectedToy = value;
                OnPropertyChanged();
                ((DelegateCommand)RemoveCommand).RaiseCanExecuteChanged();
            }
        }

        #endregion

        #region Event Handlers

        // private async void OpenDetailView(OpenDetailViewEventArgs args)
        // {
        // Int64 startTicks = Log.EVENT("(ToyDetailViewModel) Enter", Common.LOG_APPNAME);

        // await LoadAsync(args.Id);

        // Log.EVENT("(ToyDetailViewModel) Exit", Common.LOG_APPNAME, startTicks);
        // }

        #endregion

        #region Public Methods

        public override async Task LoadAsync(int id)
        {
            Int64 startTicks = Log.VIEWMODEL("(ToyDetailViewModel) Enter Id:({id})", Common.LOG_APPNAME);

            Id = id;

            foreach (var wrapper in Toys)
            {
                wrapper.PropertyChanged -= Wrapper_PropertyChanged;
            }

            Toys.Clear();

            var items = await _ToyDataService.AllAsync();

            foreach (var model in items)
            {
                var wrapper = new ToyWrapper(model);
                wrapper.PropertyChanged += Wrapper_PropertyChanged;
                Toys.Add(wrapper);
            }

            Log.VIEWMODEL("(ToyDetailViewModel) Exit", Common.LOG_APPNAME, startTicks);
        }

        void Wrapper_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (!HasChanges)
            {
                HasChanges = _ToyDataService.HasChanges();
            }

            if (e.PropertyName == nameof(ToyWrapper.HasErrors))
            {
                ((DelegateCommand)SaveCommand).RaiseCanExecuteChanged();
            }
        }

        #endregion

        #region Protected Methods

        protected override bool DeleteCanExecute()
        {
            // TODO(crhodes)
            // Why do we need this?
            return true;
        }

        protected override async void DeleteExecute()
        {
            Int64 startTicks = Log.VIEWMODEL($"(ToyDetailViewModel) Enter Id:({SelectedToy.Id})", Common.LOG_APPNAME);

            Log.VIEWMODEL("(ToyDetailViewModel) Exit", Common.LOG_APPNAME, startTicks);
        }

        protected override bool SaveCanExecute()
        {
            return HasChanges && Toys.All(p => !p.HasErrors);
        }

        protected override async void SaveExecute()
        {
            Int64 startTicks = Log.VIEWMODEL($"(ToyDetailViewModel) Enter Id:({SelectedToy.Id})", Common.LOG_APPNAME);

            try
            {
                await _ToyDataService.UpdateAsync();

                HasChanges = _ToyDataService.HasChanges();

                PublishAfterCollectionSavedEvent();
            }
            catch (Exception ex)
            {
                // while (ex.InnerException != null)
                // {
                // ex = ex.InnerException;
                // }

                MessageDialogService.ShowInfoDialog(
                    "Error while saving the Toys, " +
                    "the data will be reloaded.  Details: " + ex);
                await LoadAsync(Id);
            }

            Log.VIEWMODEL("(ToyDetailViewModel) Exit", Common.LOG_APPNAME, startTicks);
        }

        #endregion

        #region Private Methods

        void AddExecute()
        {
            Int64 startTicks = Log.VIEWMODEL("(ToyDetailViewModel) Enter", Common.LOG_APPNAME);

            var wrapper = new ToyWrapper(new Domain.Toy());
            wrapper.PropertyChanged += Wrapper_PropertyChanged;

            _ToyDataService.Add(wrapper.Model);
            Toys.Add(wrapper);

            wrapper.Name = "";  // Trigger the validation

            Log.VIEWMODEL("(ToyDetailViewModel) Exit", Common.LOG_APPNAME, startTicks);
        }

        private async void RemoveExecute()
        {
            Int64 startTicks = Log.VIEWMODEL("(ToyDetailViewModel) Enter", Common.LOG_APPNAME);

            var isReferenced =
                await _ToyDataService.IsReferencedByCatAsync(SelectedToy.Id);

            if (isReferenced)
            {
                MessageDialogService.ShowInfoDialog(
                    $"The Cat ({SelectedToy.Name})" +
                    " can't be removed;  It is referenced by at least one Cat");
                return;
            }

            SelectedToy.PropertyChanged -= Wrapper_PropertyChanged;
            _ToyDataService.Remove(SelectedToy.Model);
            Toys.Remove(SelectedToy);
            SelectedToy = null;
            HasChanges = _ToyDataService.HasChanges();

            ((DelegateCommand)SaveCommand).RaiseCanExecuteChanged();

            Log.VIEWMODEL("(ToyDetailViewModel) Exit", Common.LOG_APPNAME, startTicks);
        }

        bool RemoveCanExecute()
        {
            return SelectedToy != null;
        }

        #endregion

        #region IInstanceCount

        private static int _instanceCountVM;

        public int InstanceCountVM
        {
            get => _instanceCountVM;
            set => _instanceCountVM = value;
        }

        #endregion
    }
}
