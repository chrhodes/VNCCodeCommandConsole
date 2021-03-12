using System.Collections.Generic;
using System.Threading.Tasks;

using VNC.Core.DomainServices;

namespace VNCCodeCommandConsole.DomainServices
{
    public interface IToyLookupDataService
    {
        Task<IEnumerable<LookupItem>> GetToyLookupAsync();
    }
}
