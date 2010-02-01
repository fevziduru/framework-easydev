using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.EPS.Portal;

namespace EasyDev.EPS.BusinessObject
{
    public interface IBO
    {
        bool Save(IModel entity);

        bool Delete(IModel entity);

        bool Update(IModel entity);

        IModel FindByInstance(IModel instance);

        IList<IModel> FindAllInstance();

        string GetNextSequenceId(string tableName);
    }
}
