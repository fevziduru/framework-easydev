using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Data;
using System.Reflection;
using EasyDev.Cache;
using EasyDev.EPS.BusinessObject.GenericList;
using EasyDev.Util;

namespace EasyDev.EPS.GenericList
{
    public class GenericListService : AbstractService
    {
        #region Public Fields
        /// <summary>
        /// 列表数据
        /// </summary>
        public static readonly string LIST_DATA = "LIST_DATA";

        /// <summary>
        /// 是否是强类型数据
        /// </summary>
        public static readonly string IS_STRONG_TYPE = "IS_STRONG_TYPE";

        /// <summary>
        /// 操作是否前置
        /// </summary>
        public static readonly string IS_HEAD_OP = "IS_HEAD_OP";

        /// <summary>
        /// 编辑ACTION名称
        /// </summary>
        public static readonly string EDIT_ACTION = "EDIT_ACTION";

        /// <summary>
        /// 浏览ACTION名称
        /// </summary>
        public static readonly string DETAILS_ACTION = "DETAILS_ACTION";

        /// <summary>
        /// 删除ACTION名称
        /// </summary>
        public static readonly string DELETE_ACTION = "DELETE_ACTION";

        /// <summary>
        /// 操作所依赖的键属性名称
        /// </summary>
        public static readonly string KEYPROPERTY_NAME = "KEYPROPERTY_NAME";

        /// <summary>
        /// 模型对象类型名称
        /// </summary>
        public static readonly string MODEL_TYPENAME = "MODEL_TYPENAME";

        /// <summary>
        /// 列表表头
        /// </summary>
        public static readonly string LIST_HEADER = "LIST_HEADER";

        /// <summary>
        /// 控制器名称
        /// </summary>
        public static readonly string CONTROLLER_NAME = "CONTROLLERNAME";

        /// <summary>
        /// 用于获取数据的SQL
        /// </summary>
        public static readonly string DATA_SQL = "DATA_SQL";

        /// <summary>
        /// 是否显示操作栏
        /// </summary>
        public static readonly string IS_OP_VISIBLE = "IS_OP_VISIBLE";

        /// <summary>
        /// 弱类型数据的时候，操作所依赖的数据表名称
        /// </summary>
        private static readonly string OPTABLE = "OPTABLE"; 
        #endregion

        /// <summary>
        /// 列表模型加载后事件
        /// </summary>
        private static object listModelLoaded = new object();

        private static object dataSqlLoaded = new object();

        /// <summary>
        /// 事件列表
        /// </summary>
        protected EventHandlerList Events = new EventHandlerList();

        /// <summary>
        /// 加载数据后处理事件
        /// </summary>
        public event EventHandler<GenericListEventArgs> ListModelLoaded
        {
            add
            {
                Events.AddHandler(listModelLoaded, value);
            }
            remove
            {
                Events.RemoveHandler(listModelLoaded, value);
            }
        }

        /// <summary>
        /// 加载数据SQL后处理事件
        /// </summary>
        public event EventHandler<GenericListEventArgs> DataSQLLoaded
        {
            add
            {
                Events.AddHandler(dataSqlLoaded, value);
            }
            remove
            {
                Events.RemoveHandler(dataSqlLoaded, value);
            }
        }

        /// <summary>
        /// 从数据库读取列表配置
        /// </summary>
        /// <param name="listId"></param>
        /// <returns></returns>
        public virtual IDictionary<string, object> GetGenericListModel(string listId)
        {
            IDictionary<string, object> listModel =
                CacheManager.Retrieve(listId) as IDictionary<string, object>;

            //取得列表配置
            List list = GetBO<ListBO>().FindById(listId);
            
            //取得列表列配置
            IEnumerable<ListColumn> listColumns = GetBO<ListColumnBO>().FindListColumnsByListID(listId);
            bool isEqual = true;
            if (listModel != null)
            {
                isEqual = (list.ISOPVISIBLE.Equals("Y", StringComparison.CurrentCultureIgnoreCase) == Convert.ToBoolean(listModel[IS_OP_VISIBLE])) &&
                    (list.ISSTRONGTYPE.Equals("Y", StringComparison.CurrentCultureIgnoreCase) == Convert.ToBoolean(listModel[IS_STRONG_TYPE])) &&
                    (list.ISHEADOP.Equals("Y", StringComparison.CurrentCultureIgnoreCase) == Convert.ToBoolean(listModel[IS_HEAD_OP])) &&
                    (list.KEYPROPERTY == ConvertUtil.ToStringOrDefault(listModel[KEYPROPERTY_NAME])) &&
                    (list.EDITACTION == ConvertUtil.ToStringOrDefault(listModel[EDIT_ACTION])) &&
                    (list.DELETEACTION == ConvertUtil.ToStringOrDefault(listModel[DELETE_ACTION])) &&
                    (list.DETAILSACTION == ConvertUtil.ToStringOrDefault(listModel[DETAILS_ACTION])) &&
                    (list.MODELTYPE == ConvertUtil.ToStringOrDefault(listModel[MODEL_TYPENAME])) &&
                    (list.CONTROLLERNAME == ConvertUtil.ToStringOrDefault(listModel[CONTROLLER_NAME])) &&
                    (ConvertUtil.ToStringOrDefault(list.DATASQL) == ConvertUtil.ToStringOrDefault(list.DATASQL)) && 
                    (listColumns == (listModel[LIST_HEADER] as IEnumerable<ListColumn>));
            }

            if (isEqual == false || listModel == null)
            {
                listModel = new Dictionary<string, object>();
                listModel.Add(IS_STRONG_TYPE, list.ISSTRONGTYPE.Equals("Y",StringComparison.CurrentCultureIgnoreCase));
                listModel.Add(DATA_SQL, list.DATASQL);
                if (list.DATASERVICETYPE != null && list.DATASERVICETYPE.Length > 0)
                {
                    
                    object dataResult = null;

                    Type serviceType = Type.GetType(list.DATASERVICETYPE);
                    IListDataService dataservice = Activator.CreateInstance(serviceType) as IListDataService;

                    if (dataservice != null)
                    {
                        if (list.DATAMETHODNAME == null || list.DATAMETHODNAME.Length == 0)
                        {
                            //如果没有设置DATAMETHODNAME则调用默认的数据获取方法
                            dataResult = dataservice.DefaultDataMethod();

                            //如果是强类型数据则将数据转换为IEnumerable<IModel>,否则转换为DataTable
                            if (list.ISSTRONGTYPE.Equals("Y", StringComparison.CurrentCultureIgnoreCase))
                            {
                                listModel.Add(LIST_DATA, (dataResult as IEnumerable<IModel>));
                            }
                            else
                            {
                                listModel.Add(LIST_DATA, (dataResult as DataTable));
                                listModel.Add(OPTABLE, (dataResult as DataTable).TableName);
                            }
                        }
                        else
                        {
                            //如果指定了DATAPROVIDER则使用与DATAPROVIDER同名的无参数的方法来获取PART数据
                            MethodInfo method = serviceType.GetMethods().FirstOrDefault(p => p.Name.Equals(list.DATAMETHODNAME));
                            if (method != null)
                            {
                                dataResult = method.Invoke(dataservice, null);
                                //如果是强类型数据则将数据转换为IEnumerable<IModel>,否则转换为DataTable
                                if (list.ISSTRONGTYPE.Equals("Y", StringComparison.CurrentCultureIgnoreCase))
                                {
                                    listModel.Add(LIST_DATA, (dataResult as IEnumerable<IModel>));
                                }
                                else
                                {
                                    listModel.Add(LIST_DATA, (dataResult as DataTable));
                                    listModel.Add(OPTABLE, (dataResult as DataTable).TableName);
                                }
                            }
                            else
                            {
                                throw new GenericListServiceException(
                                    "service invalid, a data service must have a datamethodname as the method for retrieving list data.");
                            }
                        }
                    }
                }
                else
                {
                    //通过数据库中的SQL来取值
                    if (list.DATASQL != null && list.DATASQL.Length > 0)
                    {
                        string dataSql = list.DATASQL;

                        if (dataSql == null || dataSql.Length == 0)
                        {
                            dataSql = BuildSQL(list,listColumns);
                        }

                        //处理SQL参数的事件(在客户端使用的时候处理参数，如果没有参数可以不用执行这个事件)
                        if (Events[dataSqlLoaded] != null)
                        {
                            ((EventHandler<GenericListEventArgs>)Events[dataSqlLoaded])(this, new GenericListEventArgs(dataSql));
                        }

                        DataTable dtData = GetBO<ListBO>().DefaultSession.GetDataTableFromCommand(dataSql);

                        if (listModel != null && listModel.Keys.Contains(LIST_DATA))
                        {
                            listModel[LIST_DATA] = dtData;
                        }
                        else
                        {
                            listModel.Add(LIST_DATA, dtData);
                        }

                        if (listModel != null && listModel.Keys.Contains(IS_STRONG_TYPE))
                        {
                            listModel[IS_STRONG_TYPE] = false;
                        }
                        else
                        {
                            listModel.Add(IS_STRONG_TYPE, false);
                        }
                    }
                }

                listModel.Add(IS_HEAD_OP, (list.ISHEADOP != null 
                    && list.ISHEADOP.Equals("Y", StringComparison.CurrentCultureIgnoreCase)));
                listModel.Add(EDIT_ACTION, list.EDITACTION);
                listModel.Add(DETAILS_ACTION, list.DETAILSACTION);
                listModel.Add(DELETE_ACTION, list.DELETEACTION);
                listModel.Add(CONTROLLER_NAME, list.CONTROLLERNAME);
                listModel.Add(KEYPROPERTY_NAME, list.KEYPROPERTY);
                listModel.Add(MODEL_TYPENAME, list.MODELTYPE);
                
                //如果在数据库中没有找到列配置信息则通过列表配置中的模型对象的类型来创建列表头
                if (listColumns == null || listColumns.Count<ListColumn>() == 0)
                {
                    listModel.Add(LIST_HEADER, ModelFactory.CreateModel(list.MODELTYPE).GetProperties());
                }
                else
                {
                    listModel.Add(LIST_HEADER, listColumns);
                }

                listModel.Add(IS_OP_VISIBLE, (list.ISOPVISIBLE != null && list.ISOPVISIBLE.Equals("Y", StringComparison.CurrentCultureIgnoreCase)));

                //将列表模型对象缓存
                CacheManager.Store(listId, listModel);
            }

            //列表模型加载后处理事件，用于在客户端使用的时候提供一个修改列表模型数据的机会
            if (Events[listModelLoaded] != null)
            {
                ((EventHandler<GenericListEventArgs>)Events[listModelLoaded])(this, new GenericListEventArgs(listModel));
            }

            return listModel;
        }

        private string BuildSQL(List list, IEnumerable<ListColumn> listColumns)
        {
            StringBuilder sql = new StringBuilder();
            string groupClause = string.Empty;
            string orderClause = string.Empty;

            sql.Append("SELECT ");
            foreach (ListColumn item in listColumns)
            {
                sql.Append(item.ToString() + ", ");

                if (item.ISGROUP.Equals("Y", StringComparison.CurrentCultureIgnoreCase))
                {
                    groupClause += item.ALIAS + ", ";
                }

                if (item.ORDERTYPE.Equals("NONE", StringComparison.CurrentCultureIgnoreCase) == false)
                { 
                    string tmp = item.ALIAS;
                    if(tmp == null || tmp.Length == 0)
                    {
                        tmp = item.COLUMNNAME;
                    }

                    orderClause += (tmp + " " + item.ORDERTYPE + ", ");
                }
            }

            sql.Remove(sql.Length - 1, 1);
            sql.AppendFormat(" FROM {1} ",list.FROMCLAUSE);

            if (groupClause.Length > 0)
            {
                sql.AppendFormat("GROUP BY {0} ", groupClause);
            }

            if (orderClause.Length > 0)
            {
                sql.AppendFormat("ORDER BY {0} ", orderClause);
            }

            return sql.ToString();
        }
    }
}
