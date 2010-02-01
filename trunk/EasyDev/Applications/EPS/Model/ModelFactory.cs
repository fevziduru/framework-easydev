using System;
using System.Collections.Generic;
using System.Data;
using EasyDev.Util;
using System.Reflection;

namespace EasyDev.EPS
{
    public class ModelFactory
    {
        public static T CreateModel<T>() where T : IModel, new()
        {
            return new T();
        }

        public static IModel CreateModel(string modelType)
        {
            return Activator.CreateInstance(Type.GetType(modelType)) as IModel;
        }

        /// <summary>
        /// 根据数据表中的内容创建Model
        /// </summary>
        /// <typeparam name="T">Model的类型模板</typeparam>
        /// <param name="dt">包含Model数据的DataTable</param>
        /// <param name="modelType">Model的类型</param>
        /// <returns></returns>
        public static T CreateModel<T>(DataTable dt) where T : IModel, new()
        {
            return CreateModel<T>(DataSetUtil.GetFirstRowFromDataTable(dt));
        }

        /// <summary>
        /// 根据数据表中的内容创建Model(非泛型版本)
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="modelType"></param>
        /// <returns></returns>
        public static IModel CreateModel(DataTable dt, Type modelType) //where T : IModel, new()
        {
            return CreateModel(
                DataSetUtil.GetFirstRowFromDataTable(dt), modelType);
        }

        /// <summary>
        /// 创建单个Model
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dr"></param>
        /// <param name="modelType"></param>
        /// <returns></returns>
        private static T CreateModel<T>(DataRow dr) where T : IModel, new()
        {
            try
            {
                //创建一个新的MODEL实例
                T model = Activator.CreateInstance<T>();

                PropertyInfo[] properties = typeof(T).GetProperties();

                for (int i = 0; i < properties.Length; i++)
                {
                    object tmp = dr[properties[i].Name];
                    if (tmp == DBNull.Value)
                    {
                        tmp = null;
                    }
                    properties[i].SetValue(model, tmp, null);
                }

                return model;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 创建模型对象（非泛型版本）
        /// </summary>
        /// <param name="dr"></param>
        /// <param name="modelType"></param>
        /// <returns></returns>
        private static IModel CreateModel(DataRow dr, Type modelType)// where T : IModel, new()
        {
            try
            {
                //创建一个新的MODEL实例
                IModel model = Activator.CreateInstance(modelType) as IModel;

                PropertyInfo[] properties = modelType.GetProperties();

                for (int i = 0; i < properties.Length; i++)
                {
                    object tmp = dr[properties[i].Name];
                    if (tmp == DBNull.Value)
                    {
                        tmp = null;
                    }

                    properties[i].SetValue(model, tmp, null);
                }

                return model;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 批量创建Model
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dt"></param>
        /// <param name="modelType"></param>
        /// <returns></returns>
        public static IList<T> CreateModels<T>(DataTable dt) where T : IModel, new()
        {
            try
            {
                IList<T> models = new List<T>();
                foreach (DataRow row in dt.Rows)
                {
                    models.Add(CreateModel<T>(row));
                }

                return models;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public static IList<IModel> CreateModels(DataTable dt, Type modelType)
        {
            try
            {
                IList<IModel> models = new List<IModel>();
                foreach (DataRow row in dt.Rows)
                {
                    models.Add(CreateModel(row, modelType));
                }

                return models;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }
}
