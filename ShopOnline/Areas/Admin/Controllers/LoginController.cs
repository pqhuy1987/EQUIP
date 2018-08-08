using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ShopOnline.Areas.Admin.Models;
using Models;
using ShopOnline.Areas.Admin.Code;
using System.Web.Security;

namespace ShopOnline.Areas.Admin.Controllers
{
    public class LoginController : Controller
    {
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
            if (Membership.ValidateUser(model.UserName, model.Password) && ModelState.IsValid)
            {
                //SessionHelper.SetSession(new UserSession() { UserName = model.UserName });
                FormsAuthentication.SetAuthCookie(model.UserName,model.RememberMe);
                return RedirectToAction("Index", "Home");
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

    }
}
