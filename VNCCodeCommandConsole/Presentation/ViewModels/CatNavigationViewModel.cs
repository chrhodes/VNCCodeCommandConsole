using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;

using Prism.Events;

using VNC;
using VNC.Core.Events;
using VNC.Core.Mvvm;
using VNC.Core.Services;

using VNCCodeCommandConsole.DomainServices;

namespace VNCCodeCommandConsole.Presentation.ViewModels
{
    public class CatNavigationViewModel : EventViewModelBase, ICatNavigationViewModel, IInstanceCountVM
    {

        #region Constructors, Initialization, and Load

        public CatNavigationViewModel(
                ICatLookupDataService CatLookupDataService,
                IEventAggregator eventAggregator,
                IMessageDialogService messageDialogService) : base(eventAggregator, messageDialogService)
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            _CatLookupDataService = CatLookupDataService;

            InitializeViewModel();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void InitializeViewModel()
        {
            Int64 startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            InstanceCountVM++;

            Cats = new ObservableCollection<NavigationItemViewModel>();

            EventAggregator.GetEvent<AfterDetailSavedEvent>()
                .Subscribe(AfterDetailSaved);

            EventAggregator.GetEvent<AfterDetailDeletedEvent>()
                .Subscribe(AfterDetailDeleted);

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Enums


        #endregion

        #region Structures


        #endregion

        #region Fields and Properties

        private ICatLookupDataService _CatLookupDataService;

        public ObservableCollection<NavigationItemViewModel> Cats { get; private set; }

        #endregion

        #region Event Handlers

        private void AfterDetailSaved(AfterDetailSavedEventArgs args)
        {
            Int64 startTicks = Log.EVENT_HANDLER($"Enter Id:({args.Id})", Common.LOG_CATEGORY);

            switch (args.ViewModelName)
            {
                case nameof(CatDetailViewModel):
                    AfterDetailSaved(Cats, args);
                    break;

                default:
                    return;
                    //throw new System.Exception($"AfterDetailSaved(): ViewModel {args.ViewModelName} not mapped.");
            }

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void AfterDetailDeleted(AfterDetailDeletedEventArgs args)
        {
            Int64 startTicks = Log.EVENT_HANDLER($"Enter Id:({args.Id})", Common.LOG_CATEGORY);

            switch (args.ViewModelName)
            {
                case nameof(CatDetailViewModel):
                    AfterDetailDeleted(Cats, args);
                    break;

                default:
                    return;
                    //throw new System.Exception($"AfterDetailDeleted(): ViewModel {args.ViewModelName} not mapped.");
            }

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Public Methods

        public async Task LoadAsync()
        {
            Int64 startTicks = Log.VIEWMODEL("(CatNavigationViewModel) Enter", Common.LOG_CATEGORY);

            var lookupCats = await _CatLookupDataService.GetCatLookupAsync();

            Cats.Clear();

            foreach (var item in lookupCats)
            {
                Cats.Add(
                    new NavigationItemViewModel(item.Id, item.DisplayMember,
                    nameof(CatDetailViewModel),
                    EventAggregator, MessageDialogService));
            }

            Log.VIEWMODEL("(CatNavigationViewModel) Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Protected Methods


        #endregion

        #region Private Methods


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
