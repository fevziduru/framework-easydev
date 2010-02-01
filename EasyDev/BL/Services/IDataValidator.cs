using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace EasyDev.BL.Services
{
        public interface IDataValidator
        {
                IDictionary<string, object> Validate(DataSet param);
        }
}
