using System;

using Prism.Ioc;
using Prism.Modularity;
using Prism.Regions;

using Unity;

using VNC;

using VNCCodeCommandConsole.Core;
using CCC.CodeChecks.Presentation.ViewModels;
using CCC.CodeChecks.Presentation.Views;
using CCC.CodeChecks.VB.Presentation.Views;

namespace CCC.CodeChecks
{
    public class CodeChecksModule : IModule
    {
        private readonly IRegionManager _regionManager;

        // 01

        public CodeChecksModule(IRegionManager regionManager)
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            _regionManager = regionManager;

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // 02

        public void RegisterTypes(IContainerRegistry containerRegistry)
        {
            Int64 startTicks = Log.MODULE("Enter", Common.LOG_CATEGORY);

            //containerRegistry.Register<DesignChecksViewModel>();
            //containerRegistry.Register<DesignChecks>();

            containerRegistry.Register<IDesignChecksViewModel, DesignChecksViewModel>();
            containerRegistry.Register<IDesignChecks, DesignChecks>();

            containerRegistry.Register<IPerformanceChecksViewModel, PerformanceChecksViewModel>();
            containerRegistry.Register<PerformanceChecks>();

            containerRegistry.Register<QualityChecksViewModel>();
            containerRegistry.Register<QualityChecks>();

            Log.MODULE("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // 03

        public void OnInitialized(IContainerProvider containerProvider)
        {
            Int64 startTicks = Log.MODULE("Enter", Common.LOG_CATEGORY);

            // NOTE(crhodes)
            // using typeof(TYPE) calls constructor
            // using typeof(ITYPE) resolves type (see RegisterTypes)

            _regionManager.RegisterViewWithRegion(RegionNames.DesignChecksRegion, typeof(IDesignChecks));
            _regionManager.RegisterViewWithRegion(RegionNames.PerformanceChecksRegion, typeof(PerformanceChecks));
            _regionManager.RegisterViewWithRegion(RegionNames.QualityChecksRegion, typeof(QualityChecks));

            Log.MODULE("Exit", Common.LOG_CATEGORY, startTicks);
        }
    }
}
