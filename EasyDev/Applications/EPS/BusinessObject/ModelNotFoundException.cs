using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.Common;

namespace EasyDev.EPS.BusinessObject
{
    public class ModelNotFoundException : ExceptionBase,IException
    {
        public ModelNotFoundException(string msg)
            : base(msg)
        { }
    }
}
