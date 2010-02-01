using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.EPS.GenericList;
using System.Data;

namespace EasyDev.EPS.BusinessObject.GenericList
{
    public class ListColumnBO : GenericBO<ListColumn>
    {
        protected override void Initialize()
        {
            base.Initialize();
        }

        public virtual IEnumerable<ListColumn> FindListColumnsByListID(string listId)
        {
            DataTable dt = 
                DefaultSession.GetDataTableFromCommand("SELECT * FROM LISTCOLUMN WHERE LISTID=:listid", listId);
            return ModelFactory.CreateModels<ListColumn>(dt);
        }
    }
}
