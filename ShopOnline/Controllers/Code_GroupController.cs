using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Models;
using Models.Framework;

namespace ShopOnline.Controllers
{
    public class Code_GroupController : Controller
    {
        //
        // GET: /CS_tbConstructionSiteType/

        public ActionResult Index()
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                Code_GroupViewModel model = new Code_GroupViewModel();
                model.Code_Group = db.Code_Group.OrderBy(m => m.ID).ToList();
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
        public ActionResult Create(Code_GroupViewModel collection)
        {
            try
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    Code_Group obj = new Code_Group();
                    obj.Code = collection.Code_GroupSelect.Code;
                    db.Code_Group.Add(obj);
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
                Code_GroupViewModel model = new Code_GroupViewModel();

                model.Code_GroupSelect = db.Code_Group.Find(id);

                return View("Edit", model);
            }
        }

        //
        // POST: /CS_tbConstructionSiteType/Edit/5

        [HttpPost]
        public ActionResult Save(int id, Code_GroupViewModel collection)
        {
            try
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    Code_GroupViewModel model = new Code_GroupViewModel();

                    model.Code_GroupSelect = db.Code_Group.Find(id);

                    Code_Group Exsiting_Main_Job = db.Code_Group.Find(id);

                    Exsiting_Main_Job.Code = collection.Code_GroupSelect.Code;
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
                Code_GroupViewModel model = new Code_GroupViewModel();

                model.Code_GroupSelect = db.Code_Group.Find(id);

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
                    Code_GroupViewModel model = new Code_GroupViewModel();

                    Code_Group Exsiting_Main_Job = db.Code_Group.Find(id);
                    db.Code_Group.Remove(Exsiting_Main_Job);
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
