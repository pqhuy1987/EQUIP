using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Models;
using Models.Framework;


namespace ShopOnline.Areas.Admin.Controllers
{
    [Authorize]
    public class MainController : Controller
    {
        //
        // GET: /Admin/Main/

        public ActionResult Index()
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                CateloryViewModel model = new CateloryViewModel();
                model.Thiet_Bi = db.Thiet_Bis.OrderBy(m => m.ID).ToList();
                model.Catelory = db.Catelories.OrderBy(m => m.ID).ToList();

                model.ProjectAll = new List<SelectListItem>();
                var items = new List<SelectListItem>();

                foreach (var project in model.Thiet_Bi)
                {
                    items.Add(new SelectListItem()
                    {
                        Value = project.Ten_Thiet_Bi,
                        Text = project.Ten_Thiet_Bi,
                    });
                }

                model.ProjectAll = items;


                model.SelectedCatelory = null;
                model.DisplayMode = null;

                return View(model);
            }
        }

        //
        // GET: /Admin/Main/Details/5

        public ActionResult Details(int id)
        {
            return View();
        }

        //
        // GET: /Admin/Main/Create

        public ActionResult Create()
        {
            return View();
        }

        //
        // POST: /Admin/Main/Create

        [HttpPost]
        public ActionResult Create(FormCollection collection)
        {
            try
            {
                // TODO: Add insert logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        //
        // GET: /Admin/Main/Edit/5

        public ActionResult Edit(int id)
        {
            return View();
        }

        //
        // POST: /Admin/Main/Edit/5

        [HttpPost]
        public ActionResult Edit(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add update logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        //
        // GET: /Admin/Main/Delete/5

        public ActionResult Delete(int id)
        {
            return View();
        }

        //
        // POST: /Admin/Main/Delete/5

        [HttpPost]
        public ActionResult Delete(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add delete logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }
    }
}
