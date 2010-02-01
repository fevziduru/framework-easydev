using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace EasyDev.Presentation
{
    public class OperationEventArgs : EventArgs
    {
        public DataSet Data
        {
            get;
            set;
        }

        public object PrimaryKey
        {
            get;
            set;
        }

        public OperationEventArgs(DataSet _ds)
        {
            this.Data = _ds;
        }

        public OperationEventArgs(object key)
        {
            this.PrimaryKey = key;
        }
    }
}
