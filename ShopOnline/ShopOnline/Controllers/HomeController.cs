using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ShopOnline.Controllers
{
    public class HomeController : Controller
    {
        //
        // GET: /Home/

        public ActionResult Index()
        {
            var message = new ShopOnline.Models.MesssageModel();
            message.Welcome = "Chào mừng đến với Models";
            return View(message);
        }

    }
}
