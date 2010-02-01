using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using System.Reflection;
using EasyDev.EPS.BusinessObject;
using EasyDev.EPS.Portal;

namespace EasyDev.EPS.Portal
{
        public class PortalService : AbstractService
        {
                public IList<IDictionary<string, object>> RetrieveDefaultPartsData()
                {
                        try
                        {
                                SchemaBO schemaBo = GetBO<SchemaBO>();
                                ZoneBO zoneBo = GetBO<ZoneBO>();
                                IList<IDictionary<string, object>> portalLayoutModel = new List<IDictionary<string, object>>();
                                IList<Zone> zones = schemaBo.FindZonesBySchemaID(schemaBo.GetDefaultSchema().SCHEMAID);

                                foreach (Zone zone in zones)
                                {
                                        IList<Part> parts = zoneBo.FindPartsByZoneID(zone.ZONEID);
                                        //用于存放每个ZONE的信息，包括样式、标题及ZONE下的PART集合
                                        IDictionary<string, object> zoneModel = new Dictionary<string, object>();

                                        zoneModel.Add("Style", zone.ZONESTYLE);
                                        zoneModel.Add("Title", "");

                                        //用于存放ZONE中的每个PART，以名字－数据对存放
                                        IDictionary<string, object> partData = new Dictionary<string, object>();

                                        foreach (Part part in parts)
                                        {
                                                if (part.SERVICETYPE != null && part.SERVICETYPE.Length > 0)
                                                {
                                                        Type serviceType = Type.GetType(part.SERVICETYPE);
                                                        IPortalDataService dataservice = Activator.CreateInstance(serviceType) as IPortalDataService;

                                                        if (dataservice != null)
                                                        {
                                                                if (part.DATAPROVIDER == null || part.DATAPROVIDER.Length == 0)
                                                                {
                                                                        //如果没有设置DATAPROVIDER则调用默认的数据获取方法
                                                                        partData.Add("<PortalPart-" + part.PARTID + ">:" + part.PARTNAME, dataservice.DefaultDataProvider());
                                                                }
                                                                else
                                                                {
                                                                        //如果指定了DATAPROVIDER则使用与DATAPROVIDER同名的无参数的方法来获取PART数据
                                                                        MethodInfo method = serviceType.GetMethods().FirstOrDefault(p => p.Name.Equals(part.DATAPROVIDER));
                                                                        if (method != null)
                                                                        {
                                                                                partData.Add(part.PARTNAME, method.Invoke(dataservice, null));
                                                                        }
                                                                        else
                                                                        {
                                                                                throw new PortalServiceException("service invalid, a portal service must have a dataprovider as the method for retrieving portal data.");
                                                                        }
                                                                }
                                                        }
                                                        else
                                                        {
                                                                throw new PortalServiceException("service invalid, a portal service must inherit from IPortalDataService interface, and implement the interface's method.");
                                                        }
                                                }
                                                else
                                                {
                                                        //如果某个PART内部只需要显示静态内容，则不用开发相应的服务，所以在这里传NULL，在相应的PART中其MODEL为NULL。
                                                        partData.Add("<PortalPart-" + part.PARTID + ">:" + part.PARTNAME, null);
                                                }
                                        }

                                        zoneModel.Add("Parts", partData);
                                        portalLayoutModel.Add(zoneModel);
                                }

                                return portalLayoutModel;
                        }
                        catch (PortalServiceException e)
                        {
                                throw e;
                        }
                }
        }
}
