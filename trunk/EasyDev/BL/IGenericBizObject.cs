using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Collections;

namespace EasyDev.BL
{
    /// <summary>
    /// 业务对象接口
    /// </summary>
    public interface IGenericBizObject
    {
        /// <summary>
        /// 视图集合
        /// </summary>
        DataSet Views
        { get; }

        /// <summary>
        /// 模型数据，业务对象包含的数据
        /// </summary>
        DataSet BOData
        { get; set; }

        /// <summary>
        /// 数据同步，根据数据行的状态自动同步数据到数据库
        /// </summary>
        void Synchronize();

        /// <summary>
        /// 检查模型是否为空
        /// </summary>
        /// <returns></returns>
        bool IsModelEmpty();

        /// <summary>
        /// 检查模型是否为空
        /// </summary>
        /// <returns></returns>
        bool IsEmpty();

        /// <summary>
        /// 取得模型数据的第一行
        /// </summary>
        /// <returns></returns>
        DataRow GetFirstRow();

        /// <summary>
        /// 取得模型数据集的第一张数据表
        /// </summary>
        /// <returns></returns>
        DataTable GetFirstTable();

        /// <summary>
        /// 清除模型数据
        /// </summary>
        void Clear();

        /// <summary>
        /// 加载业务对象对应的所有数据记录
        /// </summary>
        /// <returns></returns>
        bool LoadAll();

        /// <summary>
        /// 根据主键加载业务对象的数据
        /// </summary>
        /// <param name="keyValue"></param>
        /// <returns></returns>
        bool LoadByPrimaryKey(object keyValue);

        /// <summary>
        /// 根据主键加载业务对象的数据(多主键情况)
        /// </summary>
        /// <param name="keyValues"></param>
        /// <returns></returns>
        bool LoadByPrimaryKeys(IDictionary<string, object> keyValues);

        /// <summary>
        /// 根据条件加载模型数据
        /// </summary>
        /// <param name="whereCond"></param>
        /// <returns></returns>
        bool LoadByCondition(string whereCond);

        /// <summary>
        /// 创建新的数据记录
        /// </summary>
        /// <returns></returns>
        DataRow Create();

        /// <summary>
        /// 根据视图名称打开视图，返回视图数据并将当前视图添加到视图集合中
        /// </summary>
        /// <param name="viewName"></param>
        /// <returns></returns>
        DataSet LoadViewByName(string viewName);

        /// <summary>
        /// 根据条件打开视图,返回视图数据并将当前视图添加到视图集合中
        /// </summary>
        /// <param name="viewName"></param>
        /// <param name="cond"></param>
        /// <returns></returns>
        DataSet LoadViewByCondition(string viewName, string cond);

        /// <summary>
        /// 从视图集合中移除视图
        /// </summary>
        /// <param name="viewName"></param>
        void RemoveFromViews(string viewName);

        /// <summary>
        /// 根据主键删除数据
        /// </summary>
        /// <param name="key"></param>
        void DeleteByPrimaryKey(object key);

        /// <summary>
        /// 根据多主键删除数据
        /// </summary>
        /// <param name="keys"></param>
        void DeleteByPrimaryKeys(IDictionary<string, object> keys);

        /// <summary>
        /// 按条件删除数据
        /// </summary>
        /// <param name="cond"></param>
        void DeleteByCondition(string cond);

        /// <summary>
        /// 从ORACLE数据库中取SEQ号
        /// </summary>
        /// <param name="tableName"></param>
        /// <returns></returns>
        string GetNextSequenceID(string tableName);
    }
}
