using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.EPS.Portal;
using EasyDev.EPS.Attributes;

namespace EasyDev.EPS.Portal
{
        [DatabaseName("Oracle")]
        public class Part : AbstractModel
        {
                /// <summary>
                /// PARTID Field 
                /// </summary>
                [KeyProperty]
                [DisplayName("部件编码")]
                public string PARTID { get; set; }

                /// <summary>
                /// ZONEID Field 
                /// </summary>
                [DisplayName("区域编码")]
                public string ZONEID { get; set; }

                /// <summary>
                /// PARTDESC Field 
                /// </summary>
                [DisplayName("部件介绍")]
                public string PARTDESC { get; set; }

                /// <summary>
                /// PARTNAME Field 
                /// </summary>
                [DisplayName("部件名称")]
                public string PARTNAME { get; set; }

                /// <summary>
                /// PARTORDER Field 
                /// </summary>
                [DisplayName("部件序号")]
                public decimal PARTORDER { get; set; }

                /// <summary>
                /// PARTTYPE Field 
                /// </summary>
                [DisplayName("部件类型")]
                public string PARTTYPE { get; set; }

                /// <summary>
                /// SERVICETYPE Field 
                /// </summary>
                [DisplayName("部件数据服务类型")]
                public string SERVICETYPE { get; set; }

                /// <summary>
                /// VOID Field 
                /// </summary>
                [DisplayName("禁用")]
                public string VOID { get; set; }

                /// <summary>
                /// CREATED Field 
                /// </summary>
                [DisplayName("创建日期")]
                public DateTime CREATED { get; set; }

                /// <summary>
                /// MODIFIED Field 
                /// </summary>
                [DisplayName("修改日期")]
                public DateTime MODIFIED { get; set; }

                /// <summary>
                /// PARTTITLE Field 
                /// </summary>
                [DisplayName("部件标题")]
                public string PARTTITLE { get; set; }

                /// <summary>
                /// DATAPROVIDER Field 
                /// </summary>
                [DisplayName("部件数据提供者")]
                public string DATAPROVIDER { get; set; }
        }
}