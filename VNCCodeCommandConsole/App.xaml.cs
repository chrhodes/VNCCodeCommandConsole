using System;
using System.IO;
using System.Threading;
using System.Windows;

using CCC.CodeChecks;
using CCC.FindSyntax;
using CCC.ModifySyntax;

using Prism.Ioc;
using Prism.Modularity;
using Prism.Regions;
using Prism.Unity;

using VNC;
using VNC.Core.Presentation.ViewModels;
using VNC.Core.Presentation.Views;
using VNC.Core.Services;

using VNCCodeCommandConsole.Presentation.Views;

namespace VNCCodeCommandConsole
{
    public partial class App : PrismApplication
    {
        #region Constructors, Initialization, and Load

        public App()
        {
            // HACK(crhodes)
            // If don't delay a bit here, the SignalR logging infrastructure does not initialize quickly enough
            // and the first few log messages are missed.
            // NB.  All are properly recored in the log file.

            Int64 startTicks = Log.CONSTRUCTOR("Initialize SignalR", Common.LOG_CATEGORY);

            Thread.Sleep(150);

            startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY, startTicks);

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // 01

        protected override void ConfigureViewModelLocator()
        {
            Int64 startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            base.ConfigureViewModelLocator();

            Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // 02

        protected override IContainerExtension CreateContainerExtension()
        {
            Int64 startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);

            return base.CreateContainerExtension();
        }

        // 03 - Create the catalog of Modules

        protected override IModuleCatalog CreateModuleCatalog()
        {
            Int64 startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);

            return base.CreateModuleCatalog();
        }

        // 04

        protected override void RegisterRequiredTypes(IContainerRegistry containerRegistry)
        {
            Int64 startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            base.RegisterRequiredTypes(containerRegistry);

            Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // 05

        protected override void RegisterTypes(IContainerRegistry containerRegistry)
        {
            Int64 startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            //containerRegistry.RegisterSingleton<ICustomerDataService, CustomerDataServiceMock>();
            //containerRegistry.RegisterSingleton<IMaterialDataService, MaterialDataServiceMock>();

            // TODO(crhodes)
            // Think this is where we switch to using the Generic Repository.
            // But how to avoid pulling knowledge of EF Context in.  Maybe Service hides details
            // of
            //containerRegistry.RegisterSingleton<IAddressDataService, AddressDataService>();
            // AddressDataService2 has a constructor that takes a CustomPoolAndSpaDbContext.

            // Common Dialogs used my most applications.

            containerRegistry.RegisterDialog<NotificationDialog, NotificationDialogViewModel>("NotificationDialog");
            containerRegistry.RegisterDialog<OkCancelDialog, OkCancelDialogViewModel>("OkCancelDialog");


            // Add the new UI elements

            Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // 06 - Removed in Prism 8.0

        //protected override void ContainerLocator()
        //{
        //    Int64 startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

        //    base.ConfigureServiceLocator();

        //    Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
        //}

        // 07 - Configure the catalog of modules
        // Modules are loaded at Startup and must be a project reference

        protected override void ConfigureModuleCatalog(IModuleCatalog moduleCatalog)
        {
            Int64 startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            //NOTE(crhodes)
            // Order matters here.  

            moduleCatalog.AddModule(typeof(CodeChecksModule));
            moduleCatalog.AddModule(typeof(FindSyntaxModule));
            moduleCatalog.AddModule(typeof(ModifySyntaxModule));
            moduleCatalog.AddModule(typeof(VNCCodeCommandConsoleModule));

            Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // 08

        protected override void ConfigureRegionAdapterMappings(RegionAdapterMappings regionAdapterMappings)
        {
            Int64 startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            base.ConfigureRegionAdapterMappings(regionAdapterMappings);

            Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // 09

        protected override void ConfigureDefaultRegionBehaviors(IRegionBehaviorFactory regionBehaviors)
        {
            Int64 startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            base.ConfigureDefaultRegionBehaviors(regionBehaviors);

            Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // 10

        protected override void RegisterFrameworkExceptionTypes()
        {
            Int64 startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            base.RegisterFrameworkExceptionTypes();

            Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // 11

        protected override Window CreateShell()
        {
            Int64 startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);

            // TODO(crhodes)
            // Pick the shell to start with.
            //return Container.Resolve<Shell>();
            return Container.Resolve<RibbonShell>();

            // NOTE(crhodes)
            // The type of view to load into the shell is handled in VNCCodeCommandConsoleModule.cs
        }

        // 12

        protected override void InitializeShell(Window shell)
        {
            Int64 startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            base.InitializeShell(shell);

            Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
        }

        // 13

        protected override void InitializeModules()
        {
            Int64 startTicks = Log.APPLICATION_INITIALIZE("Enter", Common.LOG_CATEGORY);

            base.InitializeModules();

            Log.APPLICATION_INITIALIZE("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Enums


        #endregion

        #region Structures


        #endregion

        #region Fields and Properties


        #endregion

        #region Event Handlers

        private void Application_Startup(object sender, StartupEventArgs e)
        {
            long startTicks = Log.APPLICATION_START("Enter", Common.LOG_CATEGORY);

            try
            {
                VerifyApplicationPrerequisites();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                MessageBox.Show(ex.InnerException.ToString());
            }

            Log.APPLICATION_START("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void VerifyApplicationPrerequisites()
        {
            if (! File.Exists(Common.cCONFIG_FILE))
            {
                throw new FileNotFoundException($"Cannot find {Common.cCONFIG_FILE} - Aborting");
            }

            var versionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(typeof(int).Assembly.Location);

            Common.RuntimeVersion = versionInfo.FileVersion;

            //var v = versionInfo.Comments;
            //var v1 = versionInfo.CompanyName;
            //var v2 = versionInfo.FileDescription;
            //var v3 = versionInfo.FileName;
            //var v4 = versionInfo.FileVersion;
            //var v5 = versionInfo.ProductVersion;
            //var v6 = versionInfo.ProductName;
        }
        private void Application_Activated(object sender, EventArgs e)
        {
            long startTicks = Log.APPLICATION_START("Enter", Common.LOG_CATEGORY);


            Log.APPLICATION_START("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void Application_Deactivated(object sender, EventArgs e)
        {
            long startTicks = Log.APPLICATION_END("Enter", Common.LOG_CATEGORY);


            Log.APPLICATION_END("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void Application_Exit(object sender, ExitEventArgs e)
        {
            long startTicks = Log.APPLICATION_END("Enter", Common.LOG_CATEGORY);


            Log.APPLICATION_END("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void Application_SessionEnding(object sender, SessionEndingCancelEventArgs e)
        {
            long startTicks = Log.APPLICATION_END($"Enter: ReasonSessionEnding:({e.ReasonSessionEnding})", Common.LOG_CATEGORY);


            Log.APPLICATION_END("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void Application_DispatcherUnhandledException(object sender,
            System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            Log.Error("Unexpected error occurred. Please inform the admin."
              + Environment.NewLine + e.Exception.Message, Common.LOG_CATEGORY);

            MessageBox.Show("Unexpected error occurred. Please inform the admin."
              + Environment.NewLine + e.Exception.Message, "Unexpected error");

            e.Handled = true;
        }


        #endregion

        #region Public Methods


        #endregion

        #region Protected Methods


        #endregion

        #region Private Methods


        #endregion

    }
}
