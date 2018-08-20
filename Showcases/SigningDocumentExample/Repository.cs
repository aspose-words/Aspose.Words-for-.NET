using System;
using System.Linq;
using System.Linq.Expressions;

namespace SigningDocumentExample
{
    public class Repository<T> : IRepository<T> where T : class
    {
        internal IQueryable<T> QueryableObject;

        public Repository(IQueryable<T> queryableObject)
        {
            QueryableObject = queryableObject;
        }

        public void Insert(T entity)
        {
            throw new NotImplementedException();
        }

        public void Delete(T entity)
        {
            throw new NotImplementedException();
        }

        public T FindElement(Expression<Func<T, bool>> predicate)
        {
            return QueryableObject.FirstOrDefault(predicate);
        }
    }
}
