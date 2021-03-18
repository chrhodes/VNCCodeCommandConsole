using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Threading.Tasks;

using VNC;
using VNC.Core.DomainServices;

using VNCCodeCommandConsole.Domain;

namespace VNCCodeCommandConsole.DomainServices
{
    public class CatDataServiceMock : ICatDataService
    {
        public IEnumerable<Cat> All()
        {
            Int64 startTicks = Log.DOMAINSERVICES("Enter", Common.LOG_CATEGORY);

            // TODO(crhodes)
            // Load data from real database.
            // For now just return hard coded list.

            yield return new Cat
            {
                Id = 1,
                FieldString = "FieldString",
                FieldDouble = 2.0,
                FieldInt = 23

            };

            yield return new Cat
            {
                Id = 2,
                FieldString = null,
                FieldDouble = Double.MaxValue,
                FieldInt = int.MaxValue
            };

            Log.DOMAINSERVICES("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public Task<List<Cat>> AllAsync()
        {
            throw new NotImplementedException();
        }

        public IEnumerable<Cat> AllInclude(params Expression<Func<Cat, object>>[] includeProperties)
        {
            throw new NotImplementedException();
        }

        public Task<IEnumerable<Cat>> AllIncludeAsync(params Expression<Func<Cat, object>>[] includeProperties)
        {
            Int64 startTicks = Log.DOMAINSERVICES("Enter", Common.LOG_CATEGORY);

            throw new NotImplementedException();

            Log.DOMAINSERVICES("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public IEnumerable<Cat> FindBy(Expression<Func<Cat, bool>> predicate)
        {
            Int64 startTicks = Log.DOMAINSERVICES("Enter", Common.LOG_CATEGORY);

            throw new NotImplementedException();

            Log.DOMAINSERVICES("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public Task<IEnumerable<Cat>> FindByAsync(Expression<Func<Cat, bool>> predicate)
        {
            Int64 startTicks = Log.DOMAINSERVICES("Enter", Common.LOG_CATEGORY);

            throw new NotImplementedException();

            Log.DOMAINSERVICES("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public Cat FindById(int entityId)
        {
            Int64 startTicks = Log.DOMAINSERVICES("Enter", Common.LOG_CATEGORY);

            throw new NotImplementedException();

            Log.DOMAINSERVICES("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public Task<Cat> FindByIdAsync(int entityId)
        {
            Int64 startTicks = Log.DOMAINSERVICES("Enter", Common.LOG_CATEGORY);

            throw new NotImplementedException();

            Log.DOMAINSERVICES("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public IEnumerable<Cat> FindByInclude(Expression<Func<Cat, bool>> predicate, params Expression<Func<Cat, object>>[] includeProperties)
        {
            Int64 startTicks = Log.DOMAINSERVICES("Enter", Common.LOG_CATEGORY);

            throw new NotImplementedException();

            Log.DOMAINSERVICES("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public Task<IEnumerable<Cat>> FindByIncludeAsync(Expression<Func<Cat, bool>> predicate, params Expression<Func<Cat, object>>[] includeProperties)
        {
            Int64 startTicks = Log.DOMAINSERVICES("Enter", Common.LOG_CATEGORY);

            throw new NotImplementedException();

            Log.DOMAINSERVICES("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public bool HasChanges()
        {
            Int64 startTicks = Log.DOMAINSERVICES("Enter", Common.LOG_CATEGORY);

            throw new NotImplementedException();

            Log.DOMAINSERVICES("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public void Add(Cat entity)
        {
            Int64 startTicks = Log.DOMAINSERVICES("Enter", Common.LOG_CATEGORY);

            throw new NotImplementedException();

            Log.DOMAINSERVICES("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public void Remove(Cat entity)
        {
            Int64 startTicks = Log.DOMAINSERVICES("Enter", Common.LOG_CATEGORY);

            throw new NotImplementedException();

            Log.DOMAINSERVICES("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public void RemovePhoneNumber(CatPhoneNumber model)
        {
            Int64 startTicks = Log.DOMAINSERVICES("Enter", Common.LOG_CATEGORY);

            throw new NotImplementedException();

            Log.DOMAINSERVICES("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public void Update()
        {
            Int64 startTicks = Log.DOMAINSERVICES("Enter", Common.LOG_CATEGORY);

            throw new NotImplementedException();

            Log.DOMAINSERVICES("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public Task<Cat> UpdateAsync(Cat entity)
        {
            Int64 startTicks = Log.DOMAINSERVICES("Enter", Common.LOG_CATEGORY);

            throw new NotImplementedException();

            Log.DOMAINSERVICES("Exit", Common.LOG_CATEGORY, startTicks);
        }

        public Task UpdateAsync()
        {
            Int64 startTicks = Log.DOMAINSERVICES("Enter", Common.LOG_CATEGORY);

            throw new NotImplementedException();

            Log.DOMAINSERVICES("Exit", Common.LOG_CATEGORY, startTicks);
        }
    }
}
