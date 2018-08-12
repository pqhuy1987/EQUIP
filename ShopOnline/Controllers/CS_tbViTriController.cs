using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Models;
using Models.Framework;

namespace ShopOnline.Controllers
{
    public class CS_tbViTriController : Controller
    {
        //
        // GET: /CS_tbConstructionSiteType/

        public ActionResult Index()
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                CS_tbViTriViewModel model = new CS_tbViTriViewModel();
                model.CS_tbViTri = db.CS_tbViTri.OrderBy(m => m.ID).ToList();
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
        public ActionResult Create(CS_tbViTriViewModel collection)
        {
            try
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    CS_tbViTri obj = new CS_tbViTri();
                    obj.CS_ViTri = collection.CS_tbViTriSelect.CS_ViTri;
                    db.CS_tbViTri.Add(obj);
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
                CS_tbViTriViewModel model = new CS_tbViTriViewModel();

                model.CS_tbViTriSelect = db.CS_tbViTri.Find(id);

                return View("Edit", model);
            }
        }

        //
        // POST: /CS_tbConstructionSiteType/Edit/5

        [HttpPost]
        public ActionResult Save(int id, CS_tbViTriViewModel collection)
        {
            try
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    CS_tbViTriViewModel model = new CS_tbViTriViewModel();

                    model.CS_tbViTriSelect = db.CS_tbViTri.Find(id);

                    CS_tbViTri Exsiting_Main_Job = db.CS_tbViTri.Find(id);

                    Exsiting_Main_Job.CS_ViTri = collection.CS_tbViTriSelect.CS_ViTri;
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
                CS_tbViTriViewModel model = new CS_tbViTriViewModel();

                model.CS_tbViTriSelect = db.CS_tbViTri.Find(id);

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
                    CS_tbViTriViewModel model = new CS_tbViTriViewModel();

                    CS_tbViTri Exsiting_Main_Job = db.CS_tbViTri.Find(id);
                    db.CS_tbViTri.Remove(Exsiting_Main_Job);
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
