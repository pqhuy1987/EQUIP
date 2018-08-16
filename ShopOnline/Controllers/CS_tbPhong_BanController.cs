using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Models;
using Models.Framework;

namespace ShopOnline.Controllers
{
    public class CS_tbPhong_BanController : Controller
    {
        //
        // GET: /CS_tbPhong_Ban/

        public ActionResult Index()
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                CS_tbConstructioSiteTypeViewModel model = new CS_tbConstructioSiteTypeViewModel();
                model.CS_tbPhong_Ban = db.CS_tbPhong_Ban.OrderBy(m => m.ID).ToList();
                return View(model);
            }
        }

        //
        // GET: /CS_tbPhong_Ban/Details/5

        public ActionResult Details(int id)
        {
            return View();
        }

        //
        // GET: /CS_tbPhong_Ban/Create

        public ActionResult Create()
        {
            return View("Create");
        }

        //
        // POST: /CS_tbPhong_Ban/Create

        [HttpPost]
        public ActionResult Create(CS_tbConstructioSiteTypeViewModel collection)
        {
            try
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    CS_tbPhong_Ban obj = new CS_tbPhong_Ban();
                    obj.Type = collection.CS_tbConstructionSiteType_Select.Type;
                    db.CS_tbPhong_Ban.Add(obj);
                    db.SaveChanges();

                    return RedirectToAction("Create");
                }
            }
            catch
            {
                return View();
            }
        }

        //
        // GET: /CS_tbPhong_Ban/Edit/5

        public ActionResult Edit(int id)
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                CS_tbConstructioSiteTypeViewModel model = new CS_tbConstructioSiteTypeViewModel();

                model.CS_tbConstructionSiteType_Select = db.CS_tbPhong_Ban.Find(id);

                return View("Edit",model);
            }
        }

        //
        // POST: /CS_tbPhong_Ban/Edit/5

        [HttpPost]
        public ActionResult Save(int id, CS_tbConstructioSiteTypeViewModel collection)
        {
            try
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    CS_tbConstructioSiteTypeViewModel model = new CS_tbConstructioSiteTypeViewModel();

                    model.CS_tbConstructionSiteType_Select = db.CS_tbPhong_Ban.Find(id);

                    CS_tbPhong_Ban Exsiting_Type = db.CS_tbPhong_Ban.Find(id);

                    Exsiting_Type.Type = collection.CS_tbConstructionSiteType_Select.Type;
                    db.SaveChanges();

                    return View("Edit", model);
                }
            }
            catch
            {
                return View();
            }
        }

        //
        // GET: /CS_tbPhong_Ban/Delete/5

        public ActionResult Delete(int id)
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                CS_tbConstructioSiteTypeViewModel model = new CS_tbConstructioSiteTypeViewModel();

                model.CS_tbConstructionSiteType_Select = db.CS_tbPhong_Ban.Find(id);

                return View("Delete", model);
            }
        }

        //
        // POST: /CS_tbPhong_Ban/Delete/5

        [HttpPost]
        public ActionResult Delete(int id, FormCollection collection)
        {
            try
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    CS_tbConstructioSiteTypeViewModel model = new CS_tbConstructioSiteTypeViewModel();

                    CS_tbPhong_Ban Exsiting_Type = db.CS_tbPhong_Ban.Find(id);
                    db.CS_tbPhong_Ban.Remove(Exsiting_Type);
                    db.SaveChanges();

                    return View("Finish",model);
                }
            }
            catch
            {
                return View();
            }
        }
    }
}
