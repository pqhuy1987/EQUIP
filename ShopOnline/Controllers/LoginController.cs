using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ShopOnline.Areas.Admin.Models;
using Models;
using ShopOnline.Areas.Admin.Code;
using System.Web.Security;
using System.Configuration;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
 


namespace ShopOnline.Controllers
{
    public class LoginController : Controller
    {
        HttpCookie cookie = new HttpCookie("login");
        public static string result;
        //
        // GET: /Admin/Login/

        [HttpGet]
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]

        public ActionResult Index(LoginModel model)
        {
            //var result = new AccountModel().Login(model.UserName, model.Password);
            //if (result && ModelState.IsValid)
            if (Membership.ValidateUser(model.UserName, model.Password))
            {
                //SessionHelper.SetSession(new UserSession() { UserName = model.UserName });
                FormsAuthentication.SetAuthCookie(model.UserName,model.RememberMe);

                using (var context = GetContext())
                {
                    try
                    {
                        using (var userPrinc = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, model.UserName))
                        {
                            result = userPrinc.Name;
                            cookie.Values["username"] = result;
                            cookie.Expires = DateTime.Now.AddDays(7);
                            Response.Cookies.Add(cookie);
                        }
                    }
                    catch (Exception ex)
                    {

                    }

                }
                return RedirectToAction("Index", "Thiet_Bi");
            }
            else
            {
                ModelState.AddModelError("","Tên đăng nhập không đúng hoặc mật khẩu không đúng.");
            }
            return View(model);
        }

        public ActionResult Logout()
        {
            FormsAuthentication.SignOut();
            return RedirectToAction("Index", "Login");
        }

        private static PrincipalContext GetContext()
        {
            return new PrincipalContext(ContextType.Domain, ConfigurationManager.AppSettings["DomainAccessServer"], ConfigurationManager.AppSettings["DomainAccessUser"], ConfigurationManager.AppSettings["DomainAccessPassword"]);
        }

    }
}
