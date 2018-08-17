using System;
using System.Data.SqlClient;                 //add for SqlParameter pqhuy1987
using System.ComponentModel.DataAnnotations; //add for Required pqhuy1987
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Models.Framework;

namespace Models
{
    public class LoginModel
    {
        [Required]
        public string UserName { set; get; }

        public string Password { set; get; }

        public bool RememberMe { set; get; }
    }

    public class AccountModel
    {
        private OnlineShopDbContext context = null;

        public AccountModel()
        {
            context = new OnlineShopDbContext();
        }

        public bool Login(string userName, string password)
        {
            object[] sqlParams = 
            {
                new SqlParameter("@UserName", userName),
                new SqlParameter("@Password", password),
            };
            var res = context.Database.SqlQuery<bool>("Sp_Account_Login @UserName,@Password", sqlParams).SingleOrDefault();
            return res;
        }
    }
}
