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
        // GET: /Admin/Project

        public ActionResult Index()
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                ProjectViewModel model      = new ProjectViewModel();
                model.Project = db.Projects.OrderBy(m => m.ID).ToList();

                model.SelectedProject = null;
                //model.SelectedProject.Number_Project = 100;
                return View(model);
            }
        }

        //
        // GET: /Admin/Project/Details/5

        public ActionResult Details(int id)
        {
            return View();
        }

        //
        // GET: /Admin/Project/Create

        public ActionResult Create()
        {
            return View();
        }

        //
        // POST: /Admin/Project/Create

        [HttpPost]
        public ActionResult Create(ProjectViewModel collection)
        {
            try
            {
                    using (OnlineShopDbContext db = new OnlineShopDbContext())
                    {
                        Project obj = new Project();
                        obj.Project_Name = collection.SelectedProject.Project_Name;
                        db.Projects.Add(obj);
                        db.SaveChanges();

                        ProjectViewModel model1 = new ProjectViewModel();
                        model1.Project = db.Projects.OrderByDescending(m => m.ID).ToList();
                        model1.SelectedProject = null;
                        return RedirectToAction("Index", model1);
                    }
            }
            catch
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    ProjectViewModel model1 = new ProjectViewModel();
                    model1.Project = db.Projects.OrderBy(
                            m => m.ID).ToList();
                    model1.SelectedProject = null;
                    return View("Index", model1);
                }
            }
        }

        //
        // GET: /Admin/Project/Edit/5

        public ActionResult Edit(int id)
        {
            return View();
        }

        //
        // POST: /Admin/Project/Save/5

        [HttpPost]
        public ActionResult Save(int id, ProjectViewModel collection)
        {
            try
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    Project exsiting = db.Projects.Find(id);
                    List<Catelory> exsiting_2;
                    exsiting_2 = db.Catelories.Where(i => i.Prj_Name == exsiting.Project_Name).ToList();
                    foreach (var item1 in exsiting_2)
                    {
                        item1.Prj_Name = collection.SelectedProject.Project_Name;
                    }
                    exsiting.Project_Name = collection.SelectedProject.Project_Name;
                    db.SaveChanges();

                    ProjectViewModel model1 = new ProjectViewModel();
                    model1.Project = db.Projects.OrderBy(m => m.ID).ToList();
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
                    model1.Project = db.Projects.OrderBy(
                            m => m.ID).ToList();
                    model1.SelectedProject = null;
                    return View("Index", model1);
                }
            }
        }

        //
        // GET: /Admin/Project/Delete/5

        public ActionResult Delete(int id)
        {
            return View();
        }

        //
        // POST: /Admin/Project/Delete/5

        [HttpPost]
        public ActionResult Delete(int id, FormCollection collection)
        {
            try
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    Project existing = db.Projects.Find(id);
                    db.Projects.Remove(existing);
                    db.SaveChanges();

                    ProjectViewModel model1 = new ProjectViewModel();
                    model1.Project = db.Projects.OrderBy(
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
                    model1.Project = db.Projects.OrderBy(
                            m => m.ID).ToList();
                    model1.SelectedProject = null;
                    return View("Index", model1);
                }
            }
        }
    }
}
