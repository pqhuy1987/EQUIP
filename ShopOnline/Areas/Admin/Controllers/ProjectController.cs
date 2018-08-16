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
    public class ProjectController : Controller
    {
        //
        // GET: /Admin/Thiet_Bi

        public ActionResult Index()
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                ProjectViewModel model      = new ProjectViewModel();
                model.Thiet_Bi = db.Thiet_Bis.OrderBy(m => m.ID).ToList();

                model.SelectedProject = null;
                //model.SelectedProject.Number_Project = 100;
                return View(model);
            }
        }

        //
        // GET: /Admin/Thiet_Bi/Details/5

        public ActionResult Details(int id)
        {
            return View();
        }

        //
        // GET: /Admin/Thiet_Bi/Create

        public ActionResult Create()
        {
            return View();
        }

        //
        // POST: /Admin/Thiet_Bi/Create

        [HttpPost]
        public ActionResult Create(ProjectViewModel collection)
        {
            try
            {
                    using (OnlineShopDbContext db = new OnlineShopDbContext())
                    {
                        Thiet_Bi obj = new Thiet_Bi();
                        obj.Ten_Thiet_Bi = collection.SelectedProject.Ten_Thiet_Bi;
                        db.Thiet_Bis.Add(obj);
                        db.SaveChanges();

                        ProjectViewModel model1 = new ProjectViewModel();
                        model1.Thiet_Bi = db.Thiet_Bis.OrderByDescending(m => m.ID).ToList();
                        model1.SelectedProject = null;
                        return RedirectToAction("Index", model1);
                    }
            }
            catch
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    ProjectViewModel model1 = new ProjectViewModel();
                    model1.Thiet_Bi = db.Thiet_Bis.OrderBy(
                            m => m.ID).ToList();
                    model1.SelectedProject = null;
                    return View("Index", model1);
                }
            }
        }

        //
        // GET: /Admin/Thiet_Bi/Edit/5

        public ActionResult Edit(int id)
        {
            return View();
        }

        //
        // POST: /Admin/Thiet_Bi/Save/5

        [HttpPost]
        public ActionResult Save(int id, ProjectViewModel collection)
        {
            try
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    Thiet_Bi exsiting = db.Thiet_Bis.Find(id);
                    List<Catelory> exsiting_2;
                    exsiting_2 = db.Catelories.Where(i => i.Prj_Name == exsiting.Ten_Thiet_Bi).ToList();
                    foreach (var item1 in exsiting_2)
                    {
                        item1.Prj_Name = collection.SelectedProject.Ten_Thiet_Bi;
                    }
                    exsiting.Ten_Thiet_Bi = collection.SelectedProject.Ten_Thiet_Bi;
                    db.SaveChanges();

                    ProjectViewModel model1 = new ProjectViewModel();
                    model1.Thiet_Bi = db.Thiet_Bis.OrderBy(m => m.ID).ToList();
                    model1.DisplayMode = "Add";
                    model1.SelectedProject = null;
                    return RedirectToAction("Index", model1);
                }
            }
            catch
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    ProjectViewModel model1 = new ProjectViewModel();
                    model1.Thiet_Bi = db.Thiet_Bis.OrderBy(
                            m => m.ID).ToList();
                    model1.SelectedProject = null;
                    return View("Index", model1);
                }
            }
        }

        //
        // GET: /Admin/Thiet_Bi/Delete/5

        public ActionResult Delete(int id)
        {
            return View();
        }

        //
        // POST: /Admin/Thiet_Bi/Delete/5

        [HttpPost]
        public ActionResult Delete(int id, FormCollection collection)
        {
            try
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    Thiet_Bi existing = db.Thiet_Bis.Find(id);
                    db.Thiet_Bis.Remove(existing);
                    db.SaveChanges();

                    ProjectViewModel model1 = new ProjectViewModel();
                    model1.Thiet_Bi = db.Thiet_Bis.OrderBy(
                            m => m.ID).ToList();
                    model1.SelectedProject = null;
                    return View("Index", model1);
                }
            }
            catch
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    ProjectViewModel model1 = new ProjectViewModel();
                    model1.Thiet_Bi = db.Thiet_Bis.OrderBy(
                            m => m.ID).ToList();
                    model1.SelectedProject = null;
                    return View("Index", model1);
                }
            }
        }
    }
}
