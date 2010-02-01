using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.EPS.BusinessObject;
using EasyDev.EPS.Portal;

namespace EasyDev.EPS
{
    public class ServiceFactory
    {
        public static TService CreateService<TService>()
            where TService : IService, new()
        {
            return new TService();
        }
    }
}
