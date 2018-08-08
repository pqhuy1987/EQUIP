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
    public class CateloryController : Controller
    {
        //
        // GET: /Admin/Catelory/

        public ActionResult Index()
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                CateloryViewModel model = new CateloryViewModel();
                model.Project = db.Projects.OrderBy(m => m.ID).ToList();
                model.Catelory = db.Catelories.OrderBy(m => m.ID).ToList();

                model.ProjectAll    = new List<SelectListItem>(); 
                var items           = new List<SelectListItem>();

                foreach (var project in model.Project)
                {
                    items.Add(new SelectListItem()
                    {
                        Value = project.Project_Name,
                        Text = project.Project_Name,
                    });
                }

                model.ProjectAll = items;


                model.SelectedCatelory = null;
                model.DisplayMode = null;

                return View(model);
            }
        }

        //
        // GET: /Admin/Catelory/Details/5

        public ActionResult Details(int id)
        {
            return View();
        }

        //
        // GET: /Admin/Catelory/Create

        public ActionResult Create()
        {
            return View();
        }

        //
        // POST: /Admin/Catelory/Create

        [HttpPost]
        public ActionResult Create(CateloryViewModel collection)
        {
            try
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {

                    Catelory obj            = new Catelory();

                    obj.Prj_Name            = collection.SelectedCatelory.Prj_Name;
                    obj.Unit_Name           = collection.SelectedCatelory.Unit_Name;
                    obj.Owner_Name          = collection.SelectedCatelory.Owner_Name;
                    obj.Phone_Number        = collection.SelectedCatelory.Phone_Number;
                    obj.Person_Number       = collection.SelectedCatelory.Person_Number;
                    obj.Create_Date         = DateTime.Now;
                    obj.Status              = collection.SelectedCatelory.Status;
                    obj.Rate                = collection.SelectedCatelory.Rate;
                    obj.Job                 = collection.SelectedCatelory.Job;
                    obj.Area                = collection.SelectedCatelory.Area;
                    obj.Email               = collection.SelectedCatelory.Email;

                    db.Catelories.Add(obj);
                    db.SaveChanges();

                    CateloryViewModel model = new CateloryViewModel();
                    model.Project = db.Projects.OrderBy(m => m.ID).ToList();
                    model.Catelory = db.Catelories.OrderBy(m => m.ID).ToList();

                    model.ProjectAll = new List<SelectListItem>();

                    var items = new List<SelectListItem>();

                    foreach (var project in model.Project)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = project.Project_Name,
                            Text = project.Project_Name,
                        });
                    }

                    model.ProjectAll = items;

                    model.SelectedCatelory = null;
                    model.DisplayMode = null;

                    return RedirectToAction("Index", model);
                }
            }
            catch
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    CateloryViewModel model = new CateloryViewModel();
                    model.Project = db.Projects.OrderBy(m => m.ID).ToList();
                    model.Catelory = db.Catelories.OrderBy(m => m.ID).ToList();
                    model.SelectedCatelory = null;
                    model.DisplayMode = null;
                    return View("Index", model);
                }
            }
        }

        //
        // GET: /Admin/Catelory/Edit/5

        public ActionResult Edit(int id)
        {
            return View();
        }

        //
        // POST: /Admin/Catelory/Edit/5

        [HttpPost]
        public ActionResult Edit(int id, string number, CateloryViewModel collection)
        {
            try
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    CateloryViewModel model = new CateloryViewModel();
                    model.Project = db.Projects.OrderBy(m => m.ID).ToList();
                    model.Catelory = db.Catelories.OrderBy(m => m.ID).ToList();

                    if (number == "123")
                    {
                        model.Catelory_Project = null;
                    }
                    else
                    {
                        model.Catelory_Project = db.Catelories.Where(i => i.Phone_Number == number).ToList();
                    }

                    model.ProjectAll = new List<SelectListItem>();
                    var items = new List<SelectListItem>();

                    foreach (var project in model.Project)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = project.Project_Name,
                            Text = project.Project_Name,
                        });
                    }

                    model.ProjectAll = items;

                    model.SelectedCatelory = db.Catelories.Find(id);
                    model.DisplayMode = "Edit";

                    return View("Index", model);
                }
            }
            catch
            {
                return View();
            }
        }

        //
        // POST: /Admin/Project/Save/5

        [HttpPost]
        public ActionResult Save(int id, CateloryViewModel collection)
        {
            try
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    Catelory exsiting = db.Catelories.Find(id);
                    exsiting.Prj_Name       = collection.SelectedCatelory.Prj_Name;
                    exsiting.Unit_Name      = collection.SelectedCatelory.Unit_Name;
                    exsiting.Owner_Name     = collection.SelectedCatelory.Owner_Name;
                    exsiting.Phone_Number   = collection.SelectedCatelory.Phone_Number;
                    exsiting.Person_Number  = collection.SelectedCatelory.Person_Number;
                    //exsiting.Create_Date  = DateTime.Now;
                    exsiting.Status         = collection.SelectedCatelory.Status;
                    exsiting.Area           = collection.SelectedCatelory.Area;
                    exsiting.Email          = collection.SelectedCatelory.Email;
                    exsiting.Job            = collection.SelectedCatelory.Job;
                    exsiting.Rate           = collection.SelectedCatelory.Rate;

                    db.SaveChanges();

                    CateloryViewModel model = new CateloryViewModel();
                    model.Project = db.Projects.OrderBy(m => m.ID).ToList();
                    model.Catelory = db.Catelories.OrderBy(m => m.ID).ToList();

                    model.ProjectAll    = new List<SelectListItem>();
                    var items           = new List<SelectListItem>();

                    foreach (var project in model.Project)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = project.Project_Name,
                            Text = project.Project_Name,
                        });
                    }

                    model.ProjectAll = items;

                    model.SelectedCatelory = db.Catelories.Find(id);
                    model.DisplayMode = null;

                    return RedirectToAction("Index", model);
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
        // GET: /Admin/Catelory/Delete/5

        public ActionResult Delete(int id)
        {
            return View();
        }

        //
        // POST: /Admin/Catelory/Delete/5

        [HttpPost]
        public ActionResult Delete(int id, CateloryViewModel collection)
        {
            try
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {

                    Catelory exsiting = db.Catelories.Find(id);
                    db.Catelories.Remove(exsiting);
                    db.SaveChanges();

                    CateloryViewModel model = new CateloryViewModel();
                    model.Project = db.Projects.OrderBy(m => m.ID).ToList();
                    model.Catelory = db.Catelories.OrderBy(m => m.ID).ToList();

                    model.ProjectAll = new List<SelectListItem>();
                    var items = new List<SelectListItem>();

                    foreach (var project in model.Project)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = project.Project_Name,
                            Text = project.Project_Name,
                        });
                    }

                    model.ProjectAll = items;

                    model.SelectedCatelory = db.Catelories.Find(id);
                    model.DisplayMode = null;

                    return RedirectToAction("Index", model);
                }
            }
            catch
            {
                return View();
            }
        }
    }
}
