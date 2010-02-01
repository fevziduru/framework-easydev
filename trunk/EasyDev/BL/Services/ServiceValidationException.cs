using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.Common;

namespace EasyDev.BL.Services
{
        public class ServiceValidationException :ExceptionBase, IException
        {
                public ServiceValidationException(string msg)
                        : base(msg)
                { }
        }
}
