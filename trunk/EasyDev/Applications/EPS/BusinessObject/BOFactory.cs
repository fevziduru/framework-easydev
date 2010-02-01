using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyDev.EPS.BusinessObject
{
    public class BOFactory
    {
        public static T CreateBO<T>() where T : IBO, new()
        {
            return new T();
        }

        public static IBO CreateBO(Type boType)
        {
            return Activator.CreateInstance(boType) as IBO;
        }
    }
}
