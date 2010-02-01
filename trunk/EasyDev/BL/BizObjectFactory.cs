using System;
using System.Collections.Generic;
using EasyDev.Configuration;
using System.Data;
using System.IO;
using EasyDev.PL;

namespace EasyDev.BL
{
    /// <summary>
    /// 业务对象工厂，创建通用业务对象
    /// </summary>
    public class BizObjectFactory
    {
        private PersistenceConfigManager configMgr = null;

        private string schemeDir = "";

        private GenericDBSession session = null;

        /// <summary>
        /// 用于支持MSSqlServer2005的Schema
        /// </summary>
        public string Namespace { get; set; }

        /// <summary>
        /// 应用名称，可以加在架构名称前作为架构文件的命名空间防止多个应用架构文件同名
        /// </summary>
        public string ApplicationName { get; set; }

        /// <summary>
        /// 构造器
        /// </summary>
        private BizObjectFactory(GenericDBSession session)
        {
            this.configMgr = PersistenceConfigManager.GetInstance();
            this.schemeDir = this.configMgr.SchemaDir;
            this.session = session;
        }

        /// <summary>
        /// 用于支持MSSqlServer2005的Schema
        /// </summary>
        /// <param name="_namespace"></param>
        private BizObjectFactory(string _namespace, GenericDBSession session)
        {
            this.configMgr = PersistenceConfigManager.GetInstance();
            this.schemeDir = this.configMgr.SchemaDir;
            this.Namespace = _namespace;
            this.session = session;
        }

        /// <summary>
        /// 初始化BizObjectFactory实例
        /// </summary>
        /// <returns></returns>
        public static BizObjectFactory GetInstance(GenericDBSession session)
        {
            return new BizObjectFactory(session);
        }

        /// <summary>
        /// 用于支持MSSqlServer2005的Schema
        /// </summary>
        /// <param name="_namespace"></param>
        /// <param name="session"></param>
        /// <returns></returns>
        public static BizObjectFactory GetInstance(string _namespace, GenericDBSession session)
        {
            return new BizObjectFactory(_namespace, session);
        }

        /// <summary>
        /// 根据架构信息创建数据传输对象
        /// </summary>
        /// <param name="scheme"></param>
        /// <returns></returns>
        public IGenericBizObject CreateBizObject(string schema)
        {
            return this.CreateBizObject<DataSet>(this.session, null, schema);
        }

        /// <summary>
        /// 创建通用业务对象
        /// </summary>
        /// <param name="_namespace"></param>
        /// <param name="schema"></param>
        /// <returns></returns>
        public IGenericBizObject CreateBizObject(string _namespace, string schema)
        {
            return this.CreateBizObject<DataSet>(this.session, _namespace, schema);
        }

        /// <summary>
        /// 创建强类型数据集(强类型域模型)
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="session"></param>
        /// <param name="schema"></param>
        /// <returns></returns>
        private GenericBizObject<T> CreateBizObject<T>(GenericDBSession session, string schema)
            where T : DataSet
        {
            return CreateBizObject<T>(session, this.Namespace, schema);
        }

        public GenericBizObject<T> CreateBizObject<T>(string schema)
            where T : DataSet
        {
            return CreateBizObject<T>(this.session, schema);
        }

        /// <summary>
        /// 创建通用业务对象
        /// </summary>
        /// <typeparam name="T">业务对象中的数据实体类型</typeparam>
        /// <param name="session"></param>
        /// <param name="_namespace"></param>
        /// <param name="schema"></param>
        /// <returns></returns>
        public GenericBizObject<T> CreateBizObject<T>(GenericDBSession session, string _namespace, string schema)
            where T : DataSet
        {
            GenericBizObject<T> instanceDto = null;
            DataSet dsTemp = null;

            //用于支持MSSqlServer2005的Schema,fullName表示Schema.TableName
            string fullName = string.Empty;
            try
            {
                if (string.IsNullOrEmpty(_namespace) == false)
                {
                    fullName = _namespace + "." + schema;
                }
                else
                {
                    fullName = schema;
                }

                string filePath = "";
                if (string.IsNullOrEmpty(this.ApplicationName))
                {
                    filePath = this.schemeDir + @"\" + schema + ".xsd";
                }
                else
                {
                    filePath = this.schemeDir + @"\" + this.ApplicationName + @"\" + schema + ".xsd";
                }

                if (File.Exists(filePath))
                {
                    if (typeof(T).Name == "DataSet")
                    {
                        dsTemp = new DataSet();
                        dsTemp.ReadXmlSchema(filePath);
                    }
                    else
                    {
                        dsTemp = (DataSet)Activator.CreateInstance(typeof(T));
                    }
                }
                else
                {
                    throw new PersistenceException("_missing_scheme_file");
                }

                instanceDto = new GenericBizObject<T>((T)dsTemp, session, schema);
                instanceDto.Entity = schema;
                instanceDto.FullName = fullName;
                instanceDto.Schema = filePath;
            }
            catch (PersistenceException e)
            {
                throw e;
            }

            return instanceDto;
        }
    }
}
