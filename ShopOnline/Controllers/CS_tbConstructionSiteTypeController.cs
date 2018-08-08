using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Models;
using Models.Framework;

namespace ShopOnline.Controllers
{
    public class CS_tbConstructionSiteTypeController : Controller
    {
        //
        // GET: /CS_tbConstructionSiteType/

        public ActionResult Index()
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                CS_tbConstructioSiteTypeViewModel model = new CS_tbConstructioSiteTypeViewModel();
                model.CS_tbConstructionSiteType = db.CS_tbConstructionSiteType.OrderBy(m => m.ID).ToList();
                return View(model);
            }
        }

        //
        // GET: /CS_tbConstructionSiteType/Details/5

        public ActionResult Details(int id)
        {
            return View();
        }

        //
        // GET: /CS_tbConstructionSiteType/Create

        public ActionResult Create()
        {
            return View("Create");
        }

        //
        // POST: /CS_tbConstructionSiteType/Create

        [HttpPost]
        public ActionResult Create(CS_tbConstructioSiteTypeViewModel collection)
        {
            try
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    CS_tbConstructionSiteType obj = new CS_tbConstructionSiteType();
                    obj.Type = collection.CS_tbConstructionSiteType_Select.Type;
                    db.CS_tbConstructionSiteType.Add(obj);
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
        // GET: /CS_tbConstructionSiteType/Edit/5

        public ActionResult Edit(int id)
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                CS_tbConstructioSiteTypeViewModel model = new CS_tbConstructioSiteTypeViewModel();

                model.CS_tbConstructionSiteType_Select = db.CS_tbConstructionSiteType.Find(id);

                return View("Edit",model);
            }
        }

        //
        // POST: /CS_tbConstructionSiteType/Edit/5

        [HttpPost]
        public ActionResult Save(int id, CS_tbConstructioSiteTypeViewModel collection)
        {
            try
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    CS_tbConstructioSiteTypeViewModel model = new CS_tbConstructioSiteTypeViewModel();

                    model.CS_tbConstructionSiteType_Select = db.CS_tbConstructionSiteType.Find(id);

                    CS_tbConstructionSiteType Exsiting_Type = db.CS_tbConstructionSiteType.Find(id);

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
        // GET: /CS_tbConstructionSiteType/Delete/5

        public ActionResult Delete(int id)
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                CS_tbConstructioSiteTypeViewModel model = new CS_tbConstructioSiteTypeViewModel();

                model.CS_tbConstructionSiteType_Select = db.CS_tbConstructionSiteType.Find(id);

                return View("Delete", model);
            }
        }

        //
        // POST: /CS_tbConstructionSiteType/Delete/5

        [HttpPost]
        public ActionResult Delete(int id, FormCollection collection)
        {
            try
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    CS_tbConstructioSiteTypeViewModel model = new CS_tbConstructioSiteTypeViewModel();

                    CS_tbConstructionSiteType Exsiting_Type = db.CS_tbConstructionSiteType.Find(id);
                    db.CS_tbConstructionSiteType.Remove(Exsiting_Type);
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
