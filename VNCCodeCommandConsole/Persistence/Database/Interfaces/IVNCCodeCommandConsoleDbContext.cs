using System.Data.Entity;

using VNCCodeCommandConsole.Domain;

namespace VNCCodeCommandConsole.Persistence.Data
{
    public interface IVNCCodeCommandConsoleDbContext
    {
        int SaveChanges();

        DbSet<Cat> CatsSet { get; set; }
    }
}
