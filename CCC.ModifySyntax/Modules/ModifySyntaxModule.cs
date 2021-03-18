using System;

using Prism.Ioc;
using Prism.Modularity;
using Prism.Regions;

using Unity;

using VNC;

using VNCCodeCommandConsole.Core;

using CCC.ModifySyntax.Presentation.ViewModels;
using CCC.ModifySyntax.Presentation.Views;

namespace CCC.ModifySyntax
{
    public class ModifySyntaxModule : IModule
    {
        private readonly IRegionManager _regionManager;

        // 01

        public ModifySyntaxModule(IRegionManager regionManager)
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            _regionManager = regionManager;

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // 02

        public void RegisterTypes(IContainerRegistry containerRegistry)
        {
            Int64 startTicks = Log.MODULE("Enter", Common.LOG_CATEGORY);

            containerRegistry.Register<AddCSSyntaxViewModel>();
            containerRegistry.Register<AddCSSyntax>();

            containerRegistry.Register<RemoveCSSyntaxViewModel>();
            containerRegistry.Register<RemoveCSSyntax>();

            containerRegistry.Register<RewriteCSSyntaxViewModel>();
            containerRegistry.Register<RewriteCSSyntax>();

            Log.MODULE("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // 03

        public void OnInitialized(IContainerProvider containerProvider)
        {
            Int64 startTicks = Log.MODULE("Enter", Common.LOG_CATEGORY);

            // NOTE(crhodes)
            // using typeof(TYPE) calls constructor
            // using typeof(ITYPE) resolves type (see RegisterTypes)

            _regionManager.RegisterViewWithRegion(RegionNames.AddCSSyntaxRegion, typeof(AddCSSyntax));
            _regionManager.RegisterViewWithRegion(RegionNames.AddVBSyntaxRegion, typeof(AddVBSyntax));

            _regionManager.RegisterViewWithRegion(RegionNames.RemoveCSSyntaxRegion, typeof(RemoveCSSyntax));
            _regionManager.RegisterViewWithRegion(RegionNames.RemoveVBSyntaxRegion, typeof(RemoveVBSyntax));

            _regionManager.RegisterViewWithRegion(RegionNames.RewriteCSSyntaxRegion, typeof(RewriteCSSyntax));
            _regionManager.RegisterViewWithRegion(RegionNames.RewriteVBSyntaxRegion, typeof(RewriteVBSyntax));

            Log.MODULE("Exit", Common.LOG_CATEGORY, startTicks);
        }
    }
}
