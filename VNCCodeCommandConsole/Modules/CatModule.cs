using System;

using Prism.Ioc;
using Prism.Modularity;
using Prism.Regions;

using Unity;

using VNC;

using VNCCodeCommandConsole.Core;
using VNCCodeCommandConsole.DomainServices;
using VNCCodeCommandConsole.Presentation.ViewModels;
using VNCCodeCommandConsole.Presentation.Views;

namespace VNCCodeCommandConsole
{
    public class CatModule : IModule
    {
        private readonly IRegionManager _regionManager;

        // 01

        public CatModule(IRegionManager regionManager)
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            _regionManager = regionManager;

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // 02

        public void RegisterTypes(IContainerRegistry containerRegistry)
        {
            Int64 startTicks = Log.MODULE("Enter", Common.LOG_CATEGORY);

            containerRegistry.Register<ICatMainViewModel, CatMainViewModel>();
            containerRegistry.RegisterSingleton<ICatMain, CatMain>();

            containerRegistry.Register<ICatNavigationViewModel, CatNavigationViewModel>();
            containerRegistry.RegisterSingleton<ICatNavigation, CatNavigation>();

            containerRegistry.Register<ICatDetailViewModel, CatDetailViewModel>();
            containerRegistry.RegisterSingleton<ICatDetail, CatDetail>();

            containerRegistry.RegisterSingleton<ICatLookupDataService, CatLookupDataService>();
            containerRegistry.Register<ICatDataService, CatDataService>();

            Log.MODULE("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // 03

        public void OnInitialized(IContainerProvider containerProvider)
        {
            Int64 startTicks = Log.MODULE("Enter", Common.LOG_CATEGORY);

            // NOTE(crhodes)
            // using typeof(TYPE) calls constructor
            // using typeof(ITYPE) resolves type (see RegisterTypes)

            //this loads CatMain into the Shell loaded in CreateShell() in App.Xaml.cs
            _regionManager.RegisterViewWithRegion(RegionNames.CatMainRegion, typeof(ICatMain));

            // These load into CatMain.xaml
            _regionManager.RegisterViewWithRegion(RegionNames.CatNavigationRegion, typeof(ICatNavigation));
            _regionManager.RegisterViewWithRegion(RegionNames.CatDetailRegion, typeof(ICatDetail));

            Log.MODULE("Exit", Common.LOG_CATEGORY, startTicks);
        }
    }
}
