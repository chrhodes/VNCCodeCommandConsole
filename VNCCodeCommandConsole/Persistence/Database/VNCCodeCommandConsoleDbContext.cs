using System;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Data.Entity.ModelConfiguration.Conventions;
using System.Data.Entity.Validation;
using System.Data.SqlTypes;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using VNC;
using VNC.Core.DomainServices;

using VNCCodeCommandConsole.Domain;

namespace VNCCodeCommandConsole.Persistence.Data
{
    public class VNCCodeCommandConsoleDbContext : DbContext, IVNCCodeCommandConsoleDbContext
    {
        // TODO(crhodes)
        // Add additional DbSet<TYPE> as needed.

        public DbSet<Cat> CatsSet { get; set; }
        public DbSet<CatPhoneNumber> CatPhoneNumbersSet { get; set; }

        public DbSet<Toy> ToysSet { get; set; }

        // Name of connection string in Config

        public VNCCodeCommandConsoleDbContext() : base("VNCCodeCommandConsole_DB")
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_APPNAME);
            //Database.SetInitializer(new MigrateDatabaseToLatestVersion<VNCCodeCommandConsoleDbContext, Configuration>());

            // There are four database initialization strategies

            // 1. CreateDatabaseIfNotExists (default)

            // Database.SetInitializer<VNCCodeCommandConsoleDbContext>(new CreateDatabaseIfNotExists<VNCCodeCommandConsoleDbContext>());

            // 2. DropCreateDatabaseIfModelChanges

            // Database.SetInitializer<VNCCodeCommandConsoleDbContext>(new DropCreateDatabaseIfModelChanges<VNCCodeCommandConsoleDbContext>());

            Database.SetInitializer<VNCCodeCommandConsoleDbContext>(
                new DropCreateDatabaseIfModelChanges<VNCCodeCommandConsoleDbContext>());

            // 3. DropCreateDatabaseAlways

            // Database.SetInitializer<VNCCodeCommandConsoleDbContext>(new DropCreateDatabaseAlways<VNCCodeCommandConsoleDbContext>());

            // 4. Custom DB Initializer

            //Database.SetInitializer<VNCCodeCommandConsoleDbContext>(new VNCCodeCommandConsoleDbContextDatabaseInitializer());

            // Release builds and Dependency Injection use lambda's.  Use special handling.

            Log.CONSTRUCTOR("Exit", Common.LOG_APPNAME, startTicks);
        }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            Int64 startTicks = Log.PERSISTENCE("Enter", Common.LOG_APPNAME);

            try
            {
                // TODO(crhodes)
                // This gives us too many types including the embedded ServiceType and PoolConditionReport
                // Need some way of filtering them out, otherwise have to put pointless IsDirty field in them.

                modelBuilder.Types()
                    .Configure(c => c.Ignore("IsDirty"));

            }
            catch (InvalidOperationException ex)
            {
                Log.Error(ex, Common.LOG_APPNAME);
                // Ignore
            }

            base.OnModelCreating(modelBuilder);

            // Do not pluralize table names

            modelBuilder.Conventions.Remove<PluralizingTableNameConvention>();
            //modelBuilder.HasDefaultSchema("Domain");

            // Ignore any IsDirty property on any types pulled into model.

            // Fluent API approach to constraints
            // Alternative is to apply attributes to Model Class

            //modelBuilder.Entity<Friend>()
            //    .Property(f => f.FirstName)
            //    .IsRequired()
            //    .HasMaxLength(50);

            // Could also have a general purpose approach. See class infra.

            //modelBuilder.Configurations.Add(new FriendConfiguration());

            // Alternatively can use DataAnnotations on model class.
            Log.PERSISTENCE("Exit", Common.LOG_APPNAME, startTicks);
        }

        public override int SaveChanges()
        {
            Int64 startTicks = Log.PERSISTENCE("Enter", Common.LOG_APPNAME);

            try
            {
                // Override SaveChanges so EF will enforce use of IModificationHistory changes.

                var hist = this.ChangeTracker.Entries();

                // foreach (var item in hist)
                // {
                // var isImod = item is IModificationHistory;
                // int i = 0;
                // }

                foreach (var history in this.ChangeTracker.Entries()
                  .Where(e => e.Entity is IModificationHistory
                    && (e.State == EntityState.Added || e.State == EntityState.Modified))
                   .Select(e => e.Entity as IModificationHistory)
                  )
                {
                    history.DateModified = DateTime.Now;

                    // Use DateCreated DateTime.MinValue as new flag

                    var dt = DateTime.MinValue;
                    var dt2 = SqlDateTime.MinValue;
                    var hdc = history.DateCreated;

                    //if (history.DateCreated == DateTime.MinValue)
                    //{
                    //    history.DateCreated = DateTime.Now;
                    //}

                    if (history.DateCreated == null)
                    {
                        history.DateCreated = DateTime.Now;
                    }
                }

                // Save changes to database.

                int result = base.SaveChanges();

                // Reset the IsDirty flag on any entities implementing INotificationHistory

                foreach (var history in this.ChangeTracker.Entries()
                                              .Where(e => e.Entity is IModificationHistory)
                                              .Select(e => e.Entity as IModificationHistory))
                {
                    history.IsDirty = false;
                }

                Log.PERSISTENCE($"(VNCCodeCommandConsoleDbContext) Exit ({result})", Common.LOG_APPNAME, startTicks);

                return result;

            }
            catch (DbEntityValidationException ex)
            {
                // Display some details on errors

                var sb = new StringBuilder();

                foreach (var failure in ex.EntityValidationErrors)
                {
                    sb.AppendFormat("{0} failed validation\n", failure.Entry.Entity.GetType());
                    foreach (var error in failure.ValidationErrors)
                    {
                        sb.AppendFormat("- {0} : {1}", error.PropertyName, error.ErrorMessage);
                        sb.AppendLine();
                    }
                }

                throw new DbEntityValidationException(
                    "Entity Validation Failed - errors follow:\n" +
                    sb.ToString(), ex
                    ); // Add the original exception as the innerException
            }
            catch (DbUpdateConcurrencyException ex)
            {
                var databaseValues = ex.Entries.Single().GetDatabaseValues();

                //if (databaseValues == null)
                //{
                //    MessageDialogService.ShowInfoDialog(
                //        "The entity has been deleted by another user.  Cannot continue.");
                //    PublishAfterDetailDeletedEvent(Id);
                //    return;
                //}

                //var result = MessageDialogService.ShowOkCancelDialog(
                //    "The entity has been changed by someone else." +
                //    " Click OK to save your changes anyway; Click Cancel" +
                //    " to reload the entity from the database.", "Question");

                //if (result == MessageDialogResult.OK)   // Client Wins
                //{
                //    // Update the original values with database-values
                //    var entry = ex.Entries.Single();
                //    entry.OriginalValues.SetValues(entry.GetDatabaseValues());
                //    await saveFunc();
                //}
                //else  // Database Wins
                //{
                //    // Reload entity from database
                //    await ex.Entries.Single().ReloadAsync();
                //    await LoadAsync(Id);
                //}
                throw ex;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_APPNAME);
                throw (ex);
            }

        }

        public override async Task<int> SaveChangesAsync()
        {
            Int64 startTicks = Log.PERSISTENCE("(VNCCodeCommandConsoleDbContext) Enter", Common.LOG_APPNAME);

            int result = -1;

            try
            {
                // Override SaveChanges so EF will enforce use of IModificationHistory changes.

                var hist = this.ChangeTracker.Entries();

                // foreach (var item in hist)
                // {
                // var isImod = item is IModificationHistory;
                // int i = 0;
                // }

                foreach (var history in this.ChangeTracker.Entries()
                  .Where(e => e.Entity is IModificationHistory
                    && (e.State == EntityState.Added || e.State == EntityState.Modified))
                   .Select(e => e.Entity as IModificationHistory)
                  )
                {
                    history.DateModified = DateTime.Now;

                    // Use DateCreated DateTime.MinValue as new flag

                    var dt = DateTime.MinValue;
                    var dt2 = SqlDateTime.MinValue;
                    var hdc = history.DateCreated;

                    //if (history.DateCreated == DateTime.MinValue)
                    //{
                    //    history.DateCreated = DateTime.Now;
                    //}

                    if (history.DateCreated == null)
                    {
                        history.DateCreated = DateTime.Now;
                    }
                }

                // Save changes to database.

                result = await base.SaveChangesAsync();

                // Reset the IsDirty flag on any entities implementing INotificationHistory

                foreach (var history in this.ChangeTracker.Entries()
                                              .Where(e => e.Entity is IModificationHistory)
                                              .Select(e => e.Entity as IModificationHistory))
                {
                    history.IsDirty = false;
                }

                Log.PERSISTENCE($"(VNCCodeCommandConsoleDbContext) Exit ({result})", Common.LOG_APPNAME, startTicks);

                return result;
            }
            catch (DbEntityValidationException ex)
            {
                // Display some details on errors

                var sb = new StringBuilder();

                foreach (var failure in ex.EntityValidationErrors)
                {
                    sb.AppendFormat("{0} failed validation\n", failure.Entry.Entity.GetType());
                    foreach (var error in failure.ValidationErrors)
                    {
                        sb.AppendFormat("- {0} : {1}", error.PropertyName, error.ErrorMessage);
                        sb.AppendLine();
                    }
                }

                throw new DbEntityValidationException(
                    "Entity Validation Failed - errors follow:\n" +
                    sb.ToString(), ex
                    ); // Add the original exception as the innerException
            }
            catch (DbUpdateConcurrencyException ex)
            {
                //TODO(crhodes) Maybe this should be in ViewModelBase
                // so can more sensibly use MessageDialogService
                var databaseValues = ex.Entries.Single().GetDatabaseValues();

                //if (databaseValues == null)
                //{
                //    MessageDialogService.ShowInfoDialog(
                //        "The entity has been deleted by another user.  Cannot continue.");
                //    PublishAfterDetailDeletedEvent(Id);
                //    return;
                //}

                //var result = MessageDialogService.ShowOkCancelDialog(
                //    "The entity has been changed by someone else." +
                //    " Click OK to save your changes anyway; Click Cancel" +
                //    " to reload the entity from the database.", "Question");

                //if (result == MessageDialogResult.OK)   // Client Wins
                //{
                //    // Update the original values with database-values
                //    var entry = ex.Entries.Single();
                //    entry.OriginalValues.SetValues(entry.GetDatabaseValues());
                //    await saveFunc();
                //}
                //else  // Database Wins
                //{
                //    // Reload entity from database
                //    await ex.Entries.Single().ReloadAsync();
                //    await LoadAsync(Id);
                //}

                throw (ex);
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_APPNAME);
                throw (ex);
            }

        }
    }
}
