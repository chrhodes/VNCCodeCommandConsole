using System;

using Prism.Ioc;
using Prism.Modularity;
using Prism.Regions;

using Unity;

using VNC;

using VNCCodeCommandConsole.Core;

using CCC.FindSyntax.Presentation.ViewModels;
using CCC.FindSyntax.Presentation.Views;

namespace CCC.FindSyntax
{
    public class FindSyntaxModule : IModule
    {
        private readonly IRegionManager _regionManager;

        // 01

        public FindSyntaxModule(IRegionManager regionManager)
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            _regionManager = regionManager;

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // 02

        public void RegisterTypes(IContainerRegistry containerRegistry)
        {
            Int64 startTicks = Log.MODULE("Enter", Common.LOG_CATEGORY);

            containerRegistry.Register<FindCSSyntaxViewModel>();
            containerRegistry.Register<FindCSSyntax>();

            containerRegistry.Register<FindVBSyntaxViewModel>();
            containerRegistry.Register<FindVBSyntax>();

            containerRegistry.Register<FindCSCustomViewModel>();
            containerRegistry.Register<FindCSCustom>();

            containerRegistry.Register<FindVBCustomViewModel>();
            containerRegistry.Register<FindVBCustom>();

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
            _regionManager.RegisterViewWithRegion(RegionNames.FindCSSyntaxRegion, typeof(FindCSSyntax));
            _regionManager.RegisterViewWithRegion(RegionNames.FindVBSyntaxRegion, typeof(FindVBSyntax));

            _regionManager.RegisterViewWithRegion(RegionNames.FindCSCustomRegion, typeof(FindCSCustom));
            _regionManager.RegisterViewWithRegion(RegionNames.FindVBCustomRegion, typeof(FindVBCustom));

            Log.MODULE("Exit", Common.LOG_CATEGORY, startTicks);
        }
    }
}
