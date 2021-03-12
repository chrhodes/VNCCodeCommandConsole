using VNC.Core.DomainServices;

using VNCCodeCommandConsole.Domain;

namespace VNCCodeCommandConsole.DomainServices
{
    public interface ICatDataService : IGenericRepository<Cat>
    {
        void RemovePhoneNumber(CatPhoneNumber model);
    }
}
