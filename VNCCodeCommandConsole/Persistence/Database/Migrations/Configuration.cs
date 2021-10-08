using System;
using System.Data.Entity.Migrations;

using VNC;

namespace VNCCodeCommandConsole.Persistence.Data.Migrations
{
    internal sealed class Configuration : DbMigrationsConfiguration<VNCCodeCommandConsoleDbContext>
    {
        public Configuration()
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            AutomaticMigrationsEnabled = true;

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        protected override void Seed(VNCCodeCommandConsoleDbContext context)
        {
            Int64 startTicks = Log.PERSISTENCE("Enter", Common.LOG_CATEGORY);

            //  This method will be called after migrating to the latest version.

            SeedInitialDatabaseTables(context);
            base.Seed(context);

            Log.PERSISTENCE("Exit", Common.LOG_CATEGORY, startTicks);
        }

        void SeedInitialDatabaseTables(VNCCodeCommandConsoleDbContext context)
        {
            Int64 startTicks = Log.PERSISTENCE("Enter", Common.LOG_CATEGORY);

            //  Use the DbSet<T>.AddOrUpdate() helper extension method
            //  to avoid creating duplicate seed data.

            Log.PERSISTENCE("Exit", Common.LOG_CATEGORY, startTicks);
        }
    }
}
