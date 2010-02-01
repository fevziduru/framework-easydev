using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.EPS.BusinessObject;
using System.Data;
using EasyDev.EPS.Portal;
using EasyDev.EPS;

namespace EasyDev.EPS.Security
{
        /// <summary>
        /// 创建业务对象的时候，首先要覆盖基类的Initialize方法实现模型和KeyField属性的初始化
        /// </summary>
        public class PermissionBO : GenericBO<Permissions>
        {
                protected override void Initialize()
                {
                        Model = ModelFactory.CreateModel<Permissions>();
                        KeyField = "PERMISSIONID";

                        base.Initialize();
                }

                /// <summary>
                /// 取得指定权限下的角色
                /// </summary>
                /// <param name="permissionId"></param>
                /// <returns></returns>
                public virtual IList<Roles> GetRolesOfPermission(string permissionId)
                {
                        //TODO:SQL命令通过对象化方式构建
                        string command =
                            string.Format(@"SELECT ROLES.* FROM RIGHT 
                    LEFT JOIN ROLES WHERE ROLES.ROLEID=RIGHT.ROLEID 
                    WHERE RIGHT.PERMISSIONID='{0}'", permissionId);

                        try
                        {
                                DataTable dt = DefaultSession.GetDataTableFromCommand(command);

                                return ModelFactory.CreateModels<Roles>(dt);
                        }
                        catch (Exception e)
                        {
                                throw e;
                        }
                }

                /// <summary>
                /// 保存角色授权信息
                /// </summary>
                /// <param name="roleid"></param>
                /// <param name="permissionid"></param>
                /// <returns></returns>
                public virtual bool SaveRight(string roleid, string permissionid)
                {
                        string command = string.Format(@"INSERT INTO RIGHT (ROLEID,PERMISSIONID,ACTIVED) 
                VALUES('{0}','{1}','Y')", roleid, permissionid);

                        try
                        {
                                return DefaultSession.ExecuteCommand(command) > 0;
                        }
                        catch (Exception e)
                        {
                                throw;
                        }
                }

                /// <summary>
                /// 删除角色授权信息
                /// </summary>
                /// <param name="roleid"></param>
                /// <param name="permissionid"></param>
                /// <returns></returns>
                public virtual bool DeleteRight(string roleid, string permissionid)
                {
                        string command = string.Format(@"DELETE FROM RIGHT WHERE ROLEID='{0}' AND PERMISSIONID='{1}'",
                            roleid, permissionid);

                        try
                        {
                                return DefaultSession.ExecuteCommand(command) >= 0;
                        }
                        catch (Exception e)
                        {
                                throw e;
                        }
                }

                /// <summary>
                /// 更新角色授权信息
                /// </summary>
                /// <param name="roleid"></param>
                /// <param name="permissionid"></param>
                /// <returns></returns>
                public virtual bool UpdateRight(string roleid, string permissionid)
                {
                        string command = string.Format(@"UPDATE RIGHT 
                SET ROLEID='{0}', PERMISSIONID='{1}' 
                WHERE ROLEID='{0}' AND PERMISSIONID='{1}'", roleid, permissionid);

                        try
                        {
                                return DefaultSession.ExecuteCommand(command) >= 0;
                        }
                        catch (Exception e)
                        {
                                throw e;
                        }
                }

                /// <summary>
                /// 查找角色授权信息
                /// </summary>
                /// <param name="roleid"></param>
                /// <param name="permissionId"></param>
                /// <returns></returns>
                public virtual bool FindRight(string roleid, string permissionId)
                {
                        string command = string.Format(@"SELECT Count(*) FROM RIGHT WHERE ROLEID='{0}' AND PERMISSIONID='{1}'",
                            roleid, permissionId);

                        try
                        {
                                return Convert.ToInt32(DefaultSession.GetScalarObjectFromCommand(command)) > 0;
                        }
                        catch (Exception e)
                        {
                                throw e;
                        }
                }

                /// <summary>
                /// 角色授权
                /// </summary>
                /// <param name="roleid"></param>
                /// <param name="permissionid"></param>
                /// <returns></returns>
                public virtual bool GrantRight(string roleid, string permissionid)
                {
                        try
                        {
                                if (FindRight(roleid, permissionid))
                                {
                                        return UpdateRight(roleid, permissionid);
                                }
                                else
                                {
                                        return SaveRight(roleid, permissionid);
                                }
                        }
                        catch (Exception e)
                        {
                                throw e;
                        }
                }
        }
}
