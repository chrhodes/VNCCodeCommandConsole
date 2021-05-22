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

            //containerRegistry.RegisterSingleton<ICatLookupDataService, CatLookupDataService>();
            containerRegistry.Register<IMessageDialogService, MessageDialogService>();

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
            // Order matters here.  Application depends on types in Cat
            moduleCatalog.AddModule(typeof(CatModule));
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
                // HACK(crhodes)
                // Commented all this out for now.

                //AppDomain.CurrentDomain.UnhandledException +=
                //              new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);

                //Common.CurrentUser = new WindowsPrincipal(WindowsIdentity.GetCurrent());

                //if (Data.Config.ADBypass)
                //{
                //    Common.IsAdministrator = true;
                //    Common.IsBetaUser = true;
                //}
                //else
                //{
                //    if (!Data.Config.AD_Users_AllowAll)
                //    {

                //        bool isAuthorizedUser = VNC.ActiveDirectory.Helper.CheckGroupMembership(
                //            //"maward", 
                //            Common.CurrentUser.Identity.Name,
                //            Data.Config.ADGroup_Users,
                //            Data.Config.AD_Domain);

                //        if (!isAuthorizedUser)
                //        {
                //            MessageBox.Show(string.Format("You must be a member of {0}\\{1} to run this application.",
                //                Data.Config.AD_Domain, Data.Config.ADGroup_Users));
                //            return;
                //        }
                //    }

                //    Common.IsAdministrator = VNC.ActiveDirectory.Helper.CheckDirectGroupMembership(
                //        Common.CurrentUser.Identity.Name,
                //        Data.Config.ADGroup_Administrators,
                //        Data.Config.AD_Domain);

                //    Common.IsBetaUser = VNC.ActiveDirectory.Helper.CheckDirectGroupMembership(
                //        Common.CurrentUser.Identity.Name,
                //        Data.Config.ADGroup_BetaUsers,
                //        Data.Config.AD_Domain);

                //    Common.IsDeveloper = Common.CurrentUser.Identity.Name.Contains("crhodes") ? true : false;

                //    // Next lines are for testing UI only.  Comment out for normal operation.
                //    //Common.IsAdministrator = false;   
                //    //Common.IsBetaUser = false; 
                //    //Common.IsDeveloper = false;
                //}

                // Cannot do here as the Common.ApplicationDataSet has not been loaded.  Need to move here or do later.
                // For now this is in DXRibbonWindowMain();

                //var eventMessage = "Started";
                //SQLInformation.Helper.IndicateApplicationUsage(LOG_CATEGORY, DateTime.Now, currentUser.Identity.Name, eventMessage);

                // Launch the main window.



                // TODO(crhodes)
                // Would be cool to start this with a Window specified in the App.Config file

                //User_Interface.Windows.SplashScreen _window1 = new User_Interface.Windows.SplashScreen();


                // HACK(crhodes)
                // COmmented this out for now

                //User_Interface.Windows.DXRibbonWindowMain _window1 = new User_Interface.Windows.DXRibbonWindowMain();

                //User_Interface.Windows.About _window1 = new User_Interface.Windows.About();

                //String windowArgs = string.Empty;
                // Check for arguments; if there are some build the path to the package out of the args.
                //if (args.Args.Length > 0 && args.Args[0] != null)
                //{
                //    for (int i = 0; i < args.Args.Length; ++i)
                //    {
                //        windowArgs = args.Args[i];
                //        switch (i)
                //        {
                //            case 0: // Patient Id
                //                //patientId = windowArgs;
                //                break;
                //        }
                //    }
                //}

                // HACK(crhodes)
                // Commented this out for now

                //_window1.Show();
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
            else
            {
                // TODO(crhodes)
                // Maybe some basic checks that file valid, contains what we need, etc.
            }
        }

        private void Application_SessionEnding(object sender, SessionEndingCancelEventArgs e)
        {
            long startTicks = Log.APPLICATION_END("Enter", Common.LOG_CATEGORY );

            Log.APPLICATION_END("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void Application_Exit(object sender, ExitEventArgs e)
        {
            long startTicks = Log.APPLICATION_END("Enter", Common.LOG_CATEGORY);

            Log.APPLICATION_END("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void Application_Deactivated(object sender, EventArgs e)
        {
            long startTicks = Log.APPLICATION_END("Enter", Common.LOG_CATEGORY);

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
