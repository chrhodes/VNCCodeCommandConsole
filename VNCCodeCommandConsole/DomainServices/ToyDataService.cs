using System;
using System.Data.Entity;
using System.Threading.Tasks;

using VNC;
using VNC.Core.DomainServices;

using VNCCodeCommandConsole.Domain;
using VNCCodeCommandConsole.Persistence.Data;

namespace VNCCodeCommandConsole.DomainServices
{
    public class ToyDataService : GenericEFRepository<Toy, VNCCodeCommandConsoleDbContext>, IToyDataService
    {

        #region Constructors, Initialization, and Load

        public ToyDataService(VNCCodeCommandConsoleDbContext context)
            : base(context)
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Enums


        #endregion

        #region Structures


        #endregion

        #region Fields and Properties


        #endregion

        #region Event Handlers


        #endregion

        #region Public Methods

        public async Task<bool> IsReferencedByCatAsync(int id)
        {
            Int64 startTicks = Log.DOMAINSERVICES("(ToyDataService) Enter", Common.LOG_CATEGORY);

            var result = await Context.CatsSet.AsNoTracking()
                .AnyAsync(f => f.FavoriteToyId == id);

            Log.DOMAINSERVICES("(ToyDataService) Exit", Common.LOG_CATEGORY, startTicks);

            return result;
        }

        #endregion

        #region Protected Methods


        #endregion

        #region Private Methods


        #endregion


    }
}
