using System.Data.Common;

namespace EasyDev.PL
{
    /// <summary>
    /// 数据源对象
    /// <remarks>
    /// 数据源对象通过DbProviderFactory创建通用的ADO.NET对象
    /// </remarks>
    /// </summary>
    public class DataSourceObject : IDataSourceObject
    {
        private DbProviderFactory providerFactory = null;

        private string providerName = string.Empty;

        /// <summary>
        /// 数据提供程序名称
        /// </summary>
        public string ProviderName
        {
            get { return this.providerName; }
            set { this.providerName = value; }
        }

        /// <summary>
        /// 构造方法
        /// </summary>
        /// <param name="factory"></param>
        public DataSourceObject(DbProviderFactory factory)
        {
            this.providerFactory = factory;
        }

        /// <summary>
        /// 创建通用数据连接对象
        /// </summary>
        /// <returns></returns>
        public DbConnection GetDBConnection()
        {
            return this.providerFactory.CreateConnection();
        }

        /// <summary>
        /// 创建通用参数对象
        /// </summary>
        /// <returns></returns>
        public DbParameter GetDBParameter()
        {
            return this.providerFactory.CreateParameter();
        }

        /// <summary>
        /// 创建通用命令对象
        /// </summary>
        /// <returns></returns>
        public DbCommand GetDBCommand()
        {
            return this.providerFactory.CreateCommand();
        }

        /// <summary>
        /// 创建通过命令构造器
        /// </summary>
        /// <returns></returns>
        public DbCommandBuilder GetDBCommandBuilder()
        {
            return this.providerFactory.CreateCommandBuilder();
        }

        /// <summary>
        /// 创建通用数据适配器对象
        /// </summary>
        /// <returns></returns>
        public DbDataAdapter GetDBDataAdapter()
        {
            return this.providerFactory.CreateDataAdapter();
        }
    }
}
