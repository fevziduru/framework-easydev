using System;
using System.Data;

namespace EasyDev.PL
{
        public interface IGenericDBSession
        {
                /// <summary>
                /// Á¬½Ó×Ö·û´®
                /// </summary>
                string ConnectionString { get; set; }

                string SessionID { get; set; }

                int ExecuteCommand();

                IDataReader GetDataReaderFromCommand(string sqlCmd);

                IDataReader GetDataReaderFromCommand();

                IDataReader GetDataReader();

                object GetScalarObject();

                object GetScalarObjectFromCommand(string sqlCmd);

                object GetScalarObjectFromCommand();

                DataSet GetDataSet();

                DataSet GetDataSetFromCommand(string sqlCmd);

                DataSet GetDataSetFromCommand();

                DataTable GetDataTableFromCommand();

                DataTable GetDataTableFromCommand(string sqlCmd);

                DataTable GetDataTable();

                void CloseSession();

                void OpenSession();

                //bool Update(DataSet dto);

                //bool Insert(DataSet dto);

                //bool Delete(DataSet dto);
        }
}