using System.Threading.Tasks;

using VNC.Core.DomainServices;

using VNCCodeCommandConsole.Domain;

namespace VNCCodeCommandConsole.DomainServices
{
    public interface IToyDataService : IGenericRepository<Toy>
    {
        Task<bool> IsReferencedByCatAsync(int id);
    }
}
