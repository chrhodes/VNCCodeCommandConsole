using System;
using System.Data.Entity;
using System.Threading.Tasks;

using VNC;
using VNC.Core.DomainServices;

using VNCCodeCommandConsole.Domain;
using VNCCodeCommandConsole.Persistence.Data;

namespace VNCCodeCommandConsole.DomainServices
{
    public class CatDataService : GenericEFRepository<Cat, VNCCodeCommandConsoleDbContext>, ICatDataService
    {

        #region Constructors, Initialization, and Load

        public CatDataService(VNCCodeCommandConsoleDbContext context)
            : base(context)
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_APPNAME);

            Log.CONSTRUCTOR("Exit", Common.LOG_APPNAME, startTicks);
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

        public override async Task<Cat> FindByIdAsync(int id)
        {
            Int64 startTicks = Log.DOMAINSERVICES("(CatDataService) Enter", Common.LOG_APPNAME);

            var result = await Context.CatsSet
                .Include(f => f.PhoneNumbers)
                .SingleAsync(f => f.Id == id);

            Log.DOMAINSERVICES("(CatDataService) Exit", Common.LOG_APPNAME, startTicks);

            return result;
        }

        public void RemovePhoneNumber(CatPhoneNumber model)
        {
            Int64 startTicks = Log.DOMAINSERVICES("Enter", Common.LOG_APPNAME);

            Context.CatPhoneNumbersSet.Remove(model);

            Log.DOMAINSERVICES("Exit", Common.LOG_APPNAME, startTicks);
        }


        #endregion

        #region Protected Methods


        #endregion

        #region Private Methods


        #endregion

    }
}
