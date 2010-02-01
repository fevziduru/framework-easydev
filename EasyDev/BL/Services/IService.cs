using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using EasyDev.PL;
using EasyDev.SQMS;

namespace EasyDev.BL
{
        /// <summary>
        /// 服务公共接口
        /// </summary>
        public interface IService
        {
                DataRow Create();

                string GenerateCode();

                string BOName { get; }

                string SessionName { get; set; }

                DataSet LoadByCondition(string cond);

                void DeleteByKeys(IDictionary<string, object> keys);

                void DeleteByKey(object key);

                DataSet LoadViewByName(string name);

                DataSet LoadAll();
                
                void Save(DataSet dsSave);

                DataSet LoadByKeys(IDictionary<string, object> keys, bool isCreateNew);

                DataSet LoadByKeys(IDictionary<string, object> keys);

                DataSet LoadByKey(object key, bool isCreateNew);

                DataSet LoadByKey(object key);

                string GetNextSequenceID();

                GenericDBSession GetSession(string name);

                object GetReferenceValue(string refTextField, string refTable, string refKey, string refValue);

                object GetReferenceValue(GenericDBSession session, string refTextField, string refTable, string refKey, string refValue);
        }
}
