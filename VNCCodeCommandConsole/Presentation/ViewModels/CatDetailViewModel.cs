using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Input;

using Prism.Commands;
using Prism.Events;

using VNC;
using VNC.Core.DomainServices;
using VNC.Core.Events;
using VNC.Core.Mvvm;
using VNC.Core.Services;

using VNCCodeCommandConsole.Domain;
using VNCCodeCommandConsole.DomainServices;
using VNCCodeCommandConsole.Presentation.ModelWrappers;

namespace VNCCodeCommandConsole.Presentation.ViewModels
{
    public class CatDetailViewModel : DetailViewModelBase, ICatDetailViewModel, IInstanceCountVM
    {
        #region Contructors, Initialization, and Load

        public CatDetailViewModel(
            ICatDataService CatDataService,
            IToyLookupDataService ToyLookupDataService,
            IEventAggregator eventAggregator,
            IMessageDialogService messageDialogService) : base(eventAggregator, messageDialogService)
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            _CatDataService = CatDataService;
            _ToyLookupDataService = ToyLookupDataService;

            InitializeViewModel();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void InitializeViewModel()
        {
            Int64 startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            InstanceCountVM++;

            EventAggregator.GetEvent<AfterCollectionSavedEvent>()
                .Subscribe(AfterCollectionSaved);

            AddPhoneNumberCommand = new DelegateCommand(
                AddPhoneNumberExecute);

            RemovePhoneNumberCommand = new DelegateCommand(
                RemovePhoneNumberExecute, RemovePhoneNumberCanExecute);

            Toys = new ObservableCollection<LookupItem>();
            PhoneNumbers = new ObservableCollection<CatPhoneNumberWrapper>();

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Enums

        #endregion

        #region Structures

        #endregion

        #region Fields and Properties

        private ICatDataService _CatDataService;
        private IToyLookupDataService _ToyLookupDataService;

        public ICommand AddPhoneNumberCommand { get; private set; }
        public ICommand RemovePhoneNumberCommand { get; private set; }

        private CatPhoneNumberWrapper _selectedPhoneNumber;

        public ObservableCollection<LookupItem> Toys { get; private set; }
        public ObservableCollection<CatPhoneNumberWrapper> PhoneNumbers { get; private set; }


        private CatWrapper _Cat;

        public CatWrapper Cat
        {
            get { return _Cat; }
            private set
            {
                _Cat = value;
                OnPropertyChanged();
            }
        }

        public CatPhoneNumberWrapper SelectedPhoneNumber
        {
            get { return _selectedPhoneNumber; }
            set
            {
                _selectedPhoneNumber = value;
                OnPropertyChanged();
                ((DelegateCommand)RemovePhoneNumberCommand).RaiseCanExecuteChanged();
            }
        }

        #endregion

        #region Event Handlers

        private async void OpenDetailView(OpenDetailViewEventArgs args)
        {
            Int64 startTicks = Log.EVENT("Enter", Common.LOG_CATEGORY);

            await LoadAsync(args.Id);

            Log.EVENT("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private async void AfterCollectionSaved(AfterCollectionSavedEventArgs args)
        {
            Int64 startTicks = Log.EVENT_HANDLER("(CatDetailViewModel) Enter", Common.LOG_CATEGORY);

            if (args.ViewModelName == nameof(ToyDetailViewModel))
            {
                await LoadToysLookupAsync();
            }

            Log.EVENT_HANDLER("(CatDetailViewModel) Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Public Methods

        public override async Task LoadAsync(int id)
        {
            Int64 startTicks = Log.VIEWMODEL($"(CatDetailViewModel) Enter Id:({id})", Common.LOG_CATEGORY);

            var item = id > 0
                ? await _CatDataService.FindByIdAsync(id)
                : CreateNewCat();

            Id = item.Id;

            InitializeCat(item);

            InitializeCatPhoneNumbers(item.PhoneNumbers);

            await LoadToysLookupAsync();

            Log.VIEWMODEL("(CatDetailViewModel) Exit", Common.LOG_CATEGORY, startTicks);
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
            Int64 startTicks = Log.VIEWMODEL($"(CatDetailViewModel) Enter Id:({Cat.Id})", Common.LOG_CATEGORY);

            var result = MessageDialogService.ShowOkCancelDialog(
                "Do you really want to delete the Cat?", "Question");

            if (result == MessageDialogResult.OK)
            {
                _CatDataService.Remove(Cat.Model);

                await _CatDataService.UpdateAsync();

                PublishAfterDetailDeletedEvent(Cat.Id);
            }

            Log.VIEWMODEL("(CatDetailViewModel) Exit", Common.LOG_CATEGORY, startTicks);
        }

        protected override bool SaveCanExecute()
        {
            // TODO(crhodes)
            // Check if Cat is valid or has changes
            // This enables and disables the button

            var result = Cat != null
                && !Cat.HasErrors
                && HasChanges;

            return result;

            //return true;
        }

        protected override async void SaveExecute()
        {
            Int64 startTicks = Log.VIEWMODEL("(CatDetailViewModel) Enter Id:({Cat.Id})", Common.LOG_CATEGORY);

            await _CatDataService.UpdateAsync();

            //await SaveWithOptimisticConcurrencyAsync(_CatDataService.UpdateAsync,
            //  () =>
            //  {
            //      HasChanges = _CatDataService.HasChanges();
            //      Id = Cat.Id;
            //      RaiseDetailSavedEvent(Cat.Id, $"{Cat.FieldString}");
            //  });

            HasChanges = false;
            Id = Cat.Id;

            PublishAfterDetailSavedEvent(Cat.Id, Cat.FieldString);

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void AddPhoneNumberExecute()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            var newNumber = new CatPhoneNumberWrapper(new CatPhoneNumber());
            newNumber.PropertyChanged += CatPhoneNumberWrapper_PropertyChanged;
            PhoneNumbers.Add(newNumber);
            Cat.Model.PhoneNumbers.Add(newNumber.Model);
            newNumber.Number = ""; // Trigger validation :-)

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void RemovePhoneNumberExecute()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            SelectedPhoneNumber.PropertyChanged -= CatPhoneNumberWrapper_PropertyChanged;
            //_friendRepository.RemovePhoneNumber(SelectedPhoneNumber.Model);
            PhoneNumbers.Remove(SelectedPhoneNumber);
            SelectedPhoneNumber = null;
            HasChanges = _CatDataService.HasChanges();
            ((DelegateCommand)SaveCommand).RaiseCanExecuteChanged();

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private bool RemovePhoneNumberCanExecute()
        {
            return SelectedPhoneNumber != null;
        }

        #endregion

        #region Private Methods

        private Domain.Cat CreateNewCat()
        {
            Int64 startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            var item = new Domain.Cat();
            item.FieldDate = DateTime.Now;
            item.FieldDate2 = DateTime.Now;
            _CatDataService.Add(item);

            // TODO(crhodes)
            // Need to figure out how to handle creating new.
            // This tries to push all the way to the database which complains because
            // Haven't set Required Fields (e.g. FieldString)

            // This was our attempt to use our DataService later - but that creates a context and tries to add item which
            // throws exception

            //_dataService.Insert(item);

            // This is what was in Claudius Code (NB>  Add does not call Save Changes in his code
            //_friendRepository.Add(friend);

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);

            return item;
        }

        private void InitializeCat(Cat item)
        {
            Int64 startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            Cat = new CatWrapper(item);

            Cat.PropertyChanged += (s, e) =>
            {
                if (!HasChanges)
                {
                    HasChanges = _CatDataService.HasChanges();
                }

                if (e.PropertyName == nameof(Cat.HasErrors))
                {
                    ((DelegateCommand)SaveCommand).RaiseCanExecuteChanged();
                }

                if (e.PropertyName == nameof(Cat.FieldString))
                {
                    SetTitle();
                }
            };

            ((DelegateCommand)SaveCommand).RaiseCanExecuteChanged();

            // Little trick to trigger the validation when creating new entries
            if (Cat.Id == 0)
            {
                Cat.FieldString = "";
            }

            SetTitle();

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void InitializeCatPhoneNumbers(ICollection<CatPhoneNumber> phoneNumbers)
        {
            Int64 startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            foreach (var wrapper in PhoneNumbers)
            {
                wrapper.PropertyChanged -= CatPhoneNumberWrapper_PropertyChanged;
            }

            PhoneNumbers.Clear();

            foreach (var phoneNumber in phoneNumbers)
            {
                var wrapper = new CatPhoneNumberWrapper(phoneNumber);
                PhoneNumbers.Add(wrapper);
                wrapper.PropertyChanged += CatPhoneNumberWrapper_PropertyChanged;
            }

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void CatPhoneNumberWrapper_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            Int64 startTicks = Log.EVENT("Enter", Common.LOG_CATEGORY);

            if (!HasChanges)
            {
                HasChanges = _CatDataService.HasChanges();
            }
            if (e.PropertyName == nameof(CatPhoneNumberWrapper.HasErrors))
            {
                ((DelegateCommand)SaveCommand).RaiseCanExecuteChanged();
            }

            Log.EVENT("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private async Task LoadToysLookupAsync()
        {
            Int64 startTicks = Log.VIEWMODEL("(CatDetailViewModel) Enter", Common.LOG_CATEGORY);

            Toys.Clear();

            //ProgrammingLanguages.Add(new NullLookupItem());
            Toys.Add(new NullLookupItem { DisplayMember = " - " });

            var lookup = await _ToyLookupDataService
                                .GetToyLookupAsync();

            foreach (var lookupItem in lookup)
            {
                Toys.Add(lookupItem);
            }

            Log.VIEWMODEL("(CatDetailViewModel) Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void SetTitle()
        {
            Title = $"{Cat.FieldString}";
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
