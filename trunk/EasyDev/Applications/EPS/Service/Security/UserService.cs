using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.EPS.BusinessObject;
using EasyDev.EPS.Portal;
using EasyDev.EPS;

namespace EasyDev.EPS.Security
{
    public class UserService : AbstractService
    {
        /// <summary>
        /// 服务初始化
        /// </summary>
        protected override void Initialize()
        {
            base.Initialize();
        }

        /// <summary>
        /// 创建用户
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public virtual Users CreateUser(Users user)
        {
            try
            {
                if (user.USERID == null || user.USERID.Length == 0)
                {
                    user.USERID = GetDefaultBO<Users>().GetNextSequenceId("Users");
                }

                user.CREATED = DateTime.Now;
                user.MODIFIED = DateTime.Now;
                user.ACTIVED = "Y";

                if (GetDefaultBO<Users>().Save(user))
                {
                    return user;
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
        /// 创建用户
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public virtual Users CreateUser(string username, string userpwd)
        {
            Users user = ModelFactory.CreateModel<Users>();
            user.NAME = username;
            user.PASSWORD = userpwd;
            try
            {
                return CreateUser(user);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 创建用户
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public virtual Users CreateUser(string username, string userpwd, string email)
        {
            Users user = ModelFactory.CreateModel<Users>();
            user.NAME = username;
            user.PASSWORD = userpwd;
            user.EMAIL = email;
            try
            {
                return CreateUser(user);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 创建用户
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public virtual Users CreateUser(string username, string userpwd, string email, string question, string answer)
        {
            Users user = ModelFactory.CreateModel<Users>();
            user.NAME = username;
            user.PASSWORD = userpwd;
            user.EMAIL = email;
            user.QUESTION = question;
            user.ANSWER = answer;

            try
            {
                return CreateUser(user);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 删除用户
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public virtual bool RemoveUser(Users user)
        {
            return GetDefaultBO<Users>().Delete(user);
        }

        /// <summary>
        /// 删除用户
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public virtual bool RemoveUser(string userid)
        {
            Users user = ModelFactory.CreateModel<Users>();
            user.USERID = userid;

            return RemoveUser(user);
        }

        /// <summary>
        /// 修改用户密码
        /// </summary>
        /// <param name="user"></param>
        /// <param name="oldpasswd"></param>
        /// <param name="newpasswd"></param>
        public virtual bool ChangeUserPassword(Users user, string newpasswd)
        {
            try
            {
                Users toValidUser = ValidateUser(user);
                if (toValidUser != null)
                {
                    toValidUser.PASSWORD = newpasswd;

                    return GetDefaultBO<Users>().Update(toValidUser);
                }
                else
                {
                    throw new InvalidUserException("Invalid_user");
                }
            }
            catch (InvalidUserException e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 修改用户密码,调用此方法的前提是数据库中的用户表的用户名字段应该有唯一约束
        /// </summary>
        /// <param name="username">用户名（唯一约束）</param>
        /// <param name="oldpasswd"></param>
        /// <param name="newpasswd"></param>
        public virtual bool ChangeUserPassword(string username, string oldpasswd, string newpasswd)
        {
            try
            {
                Users user = ValidateUser(username, oldpasswd);
                if (user != null)
                {
                    user.PASSWORD = newpasswd;
                    return GetDefaultBO<Users>().Update(user);
                }
                else
                {
                    throw new InvalidUserException("Invalid_user");
                }
            }
            catch (InvalidUserException e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 修改用户密码
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public virtual bool ChangeUserPassword(string username, string email, string oldpasswd, string newpasswd)
        {
            try
            {
                Users user = ValidateUser(username,oldpasswd, email);
                if (user != null)
                {
                    user.PASSWORD = newpasswd;
                    return GetDefaultBO<Users>().Update(user);
                }
                else
                {
                    throw new InvalidUserException("Invalid_user");
                }
            }
            catch (InvalidUserException e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 验证用户
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public virtual Users ValidateUser(Users user)
        {
            return GetBO<UserBO>().FindByInstance(user) as Users;
        }

        /// <summary>
        /// 验证用户
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public virtual Users ValidateUser(string username, string passwd)
        {
            Users user = ModelFactory.CreateModel<Users>();
            user.NAME = username;
            user.PASSWORD = passwd;

            return ValidateUser(user);
        }

        /// <summary>
        /// 验证用户
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public virtual Users ValidateUser(string username, string passwd, string email)
        {
            Users user = ModelFactory.CreateModel<Users>();
            user.EMAIL = email;
            user.NAME = username;
            user.PASSWORD = passwd;

            return ValidateUser(user);
        }

        /// <summary>
        /// 用户所在的组
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public virtual IList<Roles> RolesOfUser(string userid)
        {
            return GetBO<RoleBO>().GetRolesByUserID(userid);
        }
    }
}
