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
    public class WorkCountController : Controller
    {
        //
        // GET: /Admin/WorkCount/

        public ActionResult Index()
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                WorkCountViewModel model = new WorkCountViewModel();
                
                model.Thiet_Bi = db.Thiet_Bis.OrderBy(m => m.ID).ToList();
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

                var Check = model.Thiet_Bi[0].Ten_Thiet_Bi;

                model.Catelory_Project = db.Catelories.Where(i => i.Prj_Name == Check).ToList();

                model.Number_Team_2 = model.Catelory_Project.Count();

                model.WorkCount = null;
                
                return View(model);
            }
        }

        //
        // GET: /Admin/WorkCount/Details/5

        public ActionResult Details(int id)
        {
            return View();
        }

        //
        // GET: /Admin/WorkCount/Create

        public ActionResult Create()
        {
            return View();
        }

        //
        // POST: /Admin/WorkCount/Create

        [HttpPost]
        public ActionResult Create(int id, WorkCountViewModel collection)
        {
            try
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    WorkCountViewModel model = new WorkCountViewModel();
                    model.Thiet_Bi = db.Thiet_Bis.OrderBy(m => m.ID).ToList();
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

                    var Check = model.Thiet_Bi[0].Ten_Thiet_Bi;

                    model.Catelory_Project = db.Catelories.Where(i => i.Prj_Name == Check).ToList();

                    model.Number_Team_2 = model.Catelory_Project.Count();
                    model.SelectedProject = null;
                    int j =0;

                     model.WorkCount = db.WorkCounts.Where(i => i.CreateDate == collection.SelectedWorkCount.CreateDate).ToList();

                     if (model.WorkCount.Count() != 0)
                     {
                         return View("Index", model);
                     }

                    foreach (var item in collection.Count_Number)
                    {
                        WorkCount obj = new WorkCount();

                        model.WorkCount = db.WorkCounts.Where(i => i.CreateDate == collection.SelectedWorkCount.CreateDate).ToList();

                        obj.Ten_Thiet_Bi = Check;
                        obj.Unit_Name = model.Catelory_Project[j].Unit_Name;
                        obj.CreateDate = collection.SelectedWorkCount.CreateDate;
                        obj.Unit_Number = item;

                        db.WorkCounts.Add(obj);
                        db.SaveChanges();
                        j = j + 1;

                    }
                    return View("Index", model);
                }
            }
            catch
            {
                return View();
            }
        }

        //
        // GET: /Admin/WorkCount/Edit/5

        public ActionResult Edit(int id)
        {
            return View();
        }

        //
        // POST: /Admin/WorkCount/Edit/5

        [HttpPost]
        public ActionResult Edit(int id, WorkCountViewModel collection)
        {
            try
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    WorkCountViewModel model = new WorkCountViewModel();

                    model.Thiet_Bi = db.Thiet_Bis.OrderBy(m => m.ID).ToList();
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

                    var Check = model.Thiet_Bi[0].Ten_Thiet_Bi;

                    model.Catelory_Project = db.Catelories.Where(i => i.Prj_Name == Check).ToList();

                    model.Number_Team_2 = model.Catelory_Project.Count();
                    model.SelectedProject = null;

                    model.WorkCount = db.WorkCounts.Where(i => i.Ten_Thiet_Bi == collection.SelectedProject.Ten_Thiet_Bi && i.CreateDate >= collection.StartDate && i.CreateDate <= collection.EndDate).ToList();

                    var dates = new List<DateTime>();

                    for (var dt = collection.StartDate; dt <= collection.EndDate; dt = dt.AddDays(1))
                    {
                        dates.Add(dt);
                    }

                    model.SelectDate = dates;

                    return View("Index", model);
                }
            }
            catch
            {
                return View();
            }
        }

        //
        // GET: /Admin/WorkCount/Delete/5

        public ActionResult Delete(int id)
        {
            return View();
        }

        //
        // POST: /Admin/WorkCount/Delete/5

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
