using System;
using System.Linq.Expressions;

namespace SigningDocumentExample
{
    public interface IRepository<T>
    {
        void Insert(T entity);
        void Delete(T entity);
        T FindElement(Expression<Func<T, bool>> predicate);
    }
}
