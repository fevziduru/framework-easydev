using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.EPS.BusinessObject;
using System.Data;
using EasyDev.EPS.Portal;

namespace EasyDev.EPS.BusinessObject
{
    public class ZoneBO : GenericBO<Zone>
    {
        public IList<Part> FindPartsByZoneID(string zoneid)
        {
            try
            {
                DefaultSession.PrepareParameterizedCommand(
                    @"select p.* from part p left join zone z on z.zoneid=p.zoneid where z.zoneid=:zoneid", zoneid);
                DataTable dt = DefaultSession.GetDataTable();
                return ModelFactory.CreateModels<Part>(dt);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }
}
