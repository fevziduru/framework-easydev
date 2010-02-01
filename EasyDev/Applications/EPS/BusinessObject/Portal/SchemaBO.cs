using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.EPS.BusinessObject;
using System.Data;
using EasyDev.EPS.Portal;

namespace EasyDev.EPS.Portal
{
    public class SchemaBO : GenericBO<Schema>
    {
        /// <summary>
        /// 根据方案编号取得区域设置
        /// </summary>
        /// <param name="schemaId"></param>
        /// <returns></returns>
        public IList<Zone> FindZonesBySchemaID(string schemaId)
        {
            IList<Zone> zones = new List<Zone>();

            try
            {
                /*
                 * 准备参数化命令 
                 * 也可以用以下的方法：
                 * DataTable dt = DefaultSession.GetDataTableFromCommand(
                 * @"select z.* from zoneschema zs left join schema s on s.schemaid=zs.schemaid 
                 * left join zone z on z.zoneid=zs.zoneid where schemaid=:schemaid", schemaId);
                 */
                DefaultSession.PrepareParameterizedCommand(
                    @"select z.* from zoneschema zs left join schema s on s.schemaid=zs.schemaid
                left join zone z on z.zoneid=zs.zoneid where s.schemaid=:schemaid", schemaId);

                DataTable dt = DefaultSession.GetDataTableFromCommand();

                return ModelFactory.CreateModels<Zone>(dt);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 取得缺省方案
        /// </summary>
        /// <returns></returns>
        public Schema GetDefaultSchema()
        {
            try
            {
                DataTable dt = 
                    DefaultSession.GetDataTableFromCommand("select * from schema s where s.isdefault='Y'");
                return ModelFactory.CreateModel<Schema>(dt);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }
}
