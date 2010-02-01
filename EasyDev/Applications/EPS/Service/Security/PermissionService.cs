using System.Collections.Generic;
using EasyDev.EPS.Portal;
using System;
using System.Linq;
using EasyDev.EPS;

namespace EasyDev.EPS.Security
{
    public class PermissionService : AbstractService
    {
        /// <summary>
        /// 创建权限
        /// </summary>
        /// <param name="permission"></param>
        /// <returns></returns>
        public virtual Permissions CreatePermission(Permissions permission)
        {
            try
            {
                if (permission.PERMISSIONID == null
                    || permission.PERMISSIONID.Length == 0)
                {
                    permission.PERMISSIONID = GetDefaultBO<Permissions>().GetNextSequenceId("Permissions");
                }

                permission.CREATED = DateTime.Now;
                permission.MODIFIED = DateTime.Now;
                permission.ACTIVED = "Y";

                if (GetDefaultBO<Permissions>().Save(permission))
                {
                    return permission;
                }
                else
                {
                    return null;
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 创建权限
        /// </summary>
        /// <param name="permission"></param>
        /// <returns></returns>
        public virtual Permissions CreatePermission(string permissionName)
        {
            Permissions permission = ModelFactory.CreateModel<Permissions>();
            permission.PERMISSIONNAME = permissionName;

            try
            {
                return CreatePermission(permission);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 删除权限
        /// </summary>
        /// <param name="permission"></param>
        /// <returns></returns>
        public virtual bool RemovePermission(Permissions permission)
        {
            return GetDefaultBO<Permissions>().Delete(permission);
        }

        /// <summary>
        /// 删除权限
        /// </summary>
        /// <param name="permission"></param>
        /// <returns></returns>
        public virtual bool RemovePermission(string permissionId)
        {
            Permissions permission = ModelFactory.CreateModel<Permissions>();
            permission.PERMISSIONID = permissionId;

            return RemovePermission(permission);
        }

        /// <summary>
        /// 指定权限的所有角色
        /// </summary>
        /// <param name="permissionId"></param>
        /// <returns></returns>
        public virtual  IList<Roles> RolesInPermission(string permissionId)
        {
            return GetBO<PermissionBO>().GetRolesOfPermission(permissionId);
        }

        /// <summary>
        /// 查找权限是否包含指定角色
        /// </summary>
        /// <param name="roleId"></param>
        /// <param name="permission"></param>
        /// <returns></returns>
        public virtual Roles FindRoleInPermission(string roleId, Permissions permission)
        {
            try
            {
                IList<Roles> roleList = GetBO<PermissionBO>().GetRolesOfPermission(permission.PERMISSIONID);

                return (from p in roleList
                        where p.ROLEID == roleId
                        select p).First<Roles>();
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 角色授权
        /// </summary>
        /// <param name="roleId">要授权的角色ID</param>
        /// <param name="permissions">授权列表</param>
        /// <returns>授权失败的权限点，如果返回列表元素个数为0则所有权限授权成功</returns>
        public virtual IList<Permissions> Authorize(string roleId, IList<Permissions> permissions)
        {
            try
            {
                PermissionBO bo = GetBO<PermissionBO>();
                IList<Permissions> failed = new List<Permissions>();

                foreach (Permissions item in permissions)
                {
                    if (bo.GrantRight(roleId, item.PERMISSIONID) == false)
                    {
                        failed.Add(item);
                    }
                }

                return failed;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }
}
