using System;
using System.Data.Entity;

using VNC;

namespace VNCCodeCommandConsole.Persistence.Data
{
    public class VNCCodeCommandConsoleDbContextDatabaseInitializer : CreateDatabaseIfNotExists<VNCCodeCommandConsoleDbContext>
    {
        protected override void Seed(VNCCodeCommandConsoleDbContext context)
        {
            Int64 startTicks = Log.PERSISTENCE("Enter", Common.LOG_CATEGORY);

            base.Seed(context);

            Log.PERSISTENCE("Exit", Common.LOG_CATEGORY, startTicks);
        }
    }
}
