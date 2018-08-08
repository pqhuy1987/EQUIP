using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations; //add for Required pqhuy1987
using System.Linq;
using System.Web;

namespace ShopOnline.Models
{
    public class LoginModel
    {
        [Required]
        public string UserName { set; get; }

        public string Password { set; get; }

        public bool RememberMe { set; get; }
    }
}