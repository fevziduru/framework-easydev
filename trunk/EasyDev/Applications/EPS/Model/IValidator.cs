using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyDev.EPS
{
    public interface IValidator
    {
        string DoValidate(object value);
    }
}
