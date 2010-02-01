using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.Common;

namespace EasyDev.EPS.GenericList
{
    public class GenericListServiceException : ExceptionBase , IException
    {
        public GenericListServiceException(string msg)
            : base(msg)
        { }
    }
}
