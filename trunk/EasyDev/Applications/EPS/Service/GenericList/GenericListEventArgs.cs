using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyDev.EPS.GenericList
{
    public class GenericListEventArgs : EventArgs
    {
        public GenericListEventArgs(IDictionary<string, object> obj)
        { 
            ListModel = obj;
        }

        public GenericListEventArgs(string sql)
        {
            DataSQL = sql;
        }

        public IDictionary<string, object> ListModel
        {
            get;
            set;
        }

        public string DataSQL
        {
            get;
            set;
        }
    }
}
