using System.Data;

namespace EasyDev.PL
{
    public interface IDBSession
    {
        /// <summary>
        /// 缺省初始化数据源方法
        /// </summary>
        void InitDataSource();

        /// <summary>
        /// 根据配置的数据源名称初始化数据源
        /// </summary>
        /// <param name="cfgName"></param>
        void InitDataSourceByName(string cfgName);

        /// <summary>
        /// 创建命令对象
        /// </summary>
        /// <param name="sqlCmdText"></param>
        IDbCommand CreateCommand(string sqlCmdText);

        /// <summary>
        /// 更新
        /// </summary>
        /// <param name="tableName"></param>
        /// <param name="dsToBeUpdate"></param>
        /// <returns></returns>
        //bool Update(string tableName,DataSet dsToBeUpdate);

        string ConnectionString
        {
            get;
            set;
        }

        void OpenSession();

        DataSet GetDataSetFromCommand(IDbCommand command);

        void CloseSession();
    }
}