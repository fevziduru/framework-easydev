using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Security.Principal;
using System.Threading;

namespace EasyDev.Security
{
    [Serializable]
    public class UserInfo
    {
        #region 基本属性
        public string TokenID
        {
            get;
            set;
        }

        public string UserName
        {
            get;
            set;
        }

        public string UserID
        {
            get;
            set;
        }

        public string UserAlias
        {
            get;
            set;
        }

        public string UserCode
        {
            get;
            set;
        }

        public string RoleID
        {
            get;
            set;
        }

        public string RoleName
        {
            get;
            set;
        }
        #endregion

        public override bool Equals(object obj)
        {
            UserInfo userObj = obj as UserInfo;
            if (userObj != null)
            {
                return this.UserName.Equals(userObj.UserName)
                    && this.UserAlias.Equals(userObj.UserAlias)
                    && this.UserCode.Equals(userObj.UserCode)
                    && this.UserID.Equals(userObj.UserID)
                    && this.TokenID.Equals(userObj.TokenID)
                    && this.RoleName.Equals(userObj.RoleName)
                    && this.RoleID.Equals(userObj.RoleID);
            }
            else
            {
                return false;
            }
        }
    }
}
