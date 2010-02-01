using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.Common;

namespace EasyDev.EPS.BusinessObject
{
    public class ModelNullException:ExceptionBase, IException
    {
        public ModelNullException(string msg)
            : base(msg)
        { }
    }
}
