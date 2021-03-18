using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;

using VNC;
using VNC.Core.DomainServices;

using VNCCodeCommandConsole.Persistence.Data;

namespace VNCCodeCommandConsole.DomainServices
{
    public class CatLookupDataService : ICatLookupDataService
    {

        #region Constructors, Initialization, and Load

        public CatLookupDataService(Func<VNCCodeCommandConsoleDbContext> context)
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            _contextCreator = context;

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Enums


        #endregion

        #region Structures


        #endregion

        #region Fields and Properties

        private Func<VNCCodeCommandConsoleDbContext> _contextCreator;

        #endregion

        #region Event Handlers


        #endregion

        #region Public Methods

        public async Task<IEnumerable<LookupItem>> GetCatLookupAsync()
        {
            Int64 startTicks = Log.DOMAINSERVICES("(CatLookupDataService) Enter", Common.LOG_CATEGORY);

            IEnumerable<LookupItem> result;

            using (var ctx = _contextCreator())
            {
                result = await ctx.CatsSet.AsNoTracking()
                  .Select(f =>
                  new LookupItem
                  {
                      Id = f.Id,
                      DisplayMember = f.FieldString
                  })
                  .ToListAsync();
            }

            Log.DOMAINSERVICES("(CatLookupDataService) Exit", Common.LOG_CATEGORY, startTicks);

            return result;
        }

        #endregion

        #region Protected Methods


        #endregion

        #region Private Methods


        #endregion

    }
}
