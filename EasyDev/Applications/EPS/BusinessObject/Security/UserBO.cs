using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.EPS.BusinessObject;
using EasyDev.EPS.Portal;
using System.Data;
using EasyDev.EPS;
using System.Linq.Expressions;

namespace EasyDev.EPS.Security
{
        /// <summary>
        /// 创建业务对象的时候，首先要覆盖基类的Initialize方法实现模型和KeyField属性的初始化
        /// </summary>
        public class UserBO : GenericBO<Users>
        {
                protected override void Initialize()
                {
                        Model = ModelFactory.CreateModel<Users>();
                        KeyField = "USERID";

                        base.Initialize();
                }

                /// <summary>
                /// 根据角色ID取得属于指定角色的所有用户
                /// </summary>
                /// <param name="roleId"></param>
                /// <returns></returns>
                public virtual IList<Users> GetUsersByRoleID(string roleId)
                {
                        try
                        {
                                DataTable dt = DefaultSession.GetDataTableFromCommand(@"SELECT USERS.* 
                                                  FROM USERINROLE 
                                                  LEFT JOIN USERS ON USERINROLE.USERID=USERS.USERID WHERE USERINROLE.ROLEID=:roleid", roleId);

                                return ModelFactory.CreateModels<Users>(dt);
                        }
                        catch (Exception e)
                        {
                                throw e;
                        }
                }
        }
}
