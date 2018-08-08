using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Models;
using Models.Framework;

namespace ShopOnline.Controllers
{
    public class CS_tbWorkTypeMainController : Controller
    {
        //
        // GET: /CS_tbConstructionSiteType/

        public ActionResult Index()
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                CS_tbWorkTypeMainViewModel model = new CS_tbWorkTypeMainViewModel();
                model.CS_tbWorkTypeMain = db.CS_tbWorkTypeMain.OrderBy(m => m.ID).ToList();
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
        public ActionResult Create(CS_tbWorkTypeMainViewModel collection)
        {
            try
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    CS_tbWorkTypeMain obj = new CS_tbWorkTypeMain();
                    obj.CS_WorkTypeMain = collection.CS_tbWorkTypeMainSelect.CS_WorkTypeMain;
                    db.CS_tbWorkTypeMain.Add(obj);
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
                CS_tbWorkTypeMainViewModel model = new CS_tbWorkTypeMainViewModel();

                model.CS_tbWorkTypeMainSelect = db.CS_tbWorkTypeMain.Find(id);

                return View("Edit", model);
            }
        }

        //
        // POST: /CS_tbConstructionSiteType/Edit/5

        [HttpPost]
        public ActionResult Save(int id, CS_tbWorkTypeMainViewModel collection)
        {
            try
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    CS_tbWorkTypeMainViewModel model = new CS_tbWorkTypeMainViewModel();

                    model.CS_tbWorkTypeMainSelect = db.CS_tbWorkTypeMain.Find(id);

                    CS_tbWorkTypeMain Exsiting_Main_Job = db.CS_tbWorkTypeMain.Find(id);

                    Exsiting_Main_Job.CS_WorkTypeMain = collection.CS_tbWorkTypeMainSelect.CS_WorkTypeMain;
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
                CS_tbWorkTypeMainViewModel model = new CS_tbWorkTypeMainViewModel();

                model.CS_tbWorkTypeMainSelect = db.CS_tbWorkTypeMain.Find(id);

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
                    CS_tbWorkTypeMainViewModel model = new CS_tbWorkTypeMainViewModel();

                    CS_tbWorkTypeMain Exsiting_Main_Job = db.CS_tbWorkTypeMain.Find(id);
                    db.CS_tbWorkTypeMain.Remove(Exsiting_Main_Job);
                    db.SaveChanges();

                    return View("Finish", model);
                }
            }
            catch
            {
                return View();
            }
        }
    }
}
