using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Models;
using Models.Framework;


namespace ShopOnline.Controllers
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
                    int j = 0;

                     model.WorkCount = db.WorkCounts.Where(i => i.CreateDate == collection.SelectedWorkCount.CreateDate).ToList();

                     if (model.WorkCount.Count() != 0)
                     {
                         return View("Index", model);
                     }

                    foreach (var item in collection.Count_Number)
                    {
                        WorkCount obj       = new WorkCount();
                        WorkCount obj_temp  = new WorkCount();
                        List<DateTime?>  Date_Temp    = new List<DateTime?>();

                        int Count_Temp;

                        model.WorkCount = db.WorkCounts.Where(i => i.CreateDate == collection.SelectedWorkCount.CreateDate).ToList();

                        obj.Ten_Thiet_Bi = Check;
                        obj.Unit_Name = model.Catelory_Project[j].Unit_Name;
                        obj.Unit_Job = model.Catelory_Project[j].Job;

                        Count_Temp = db.WorkCounts.Where(i => i.Unit_Name == obj.Unit_Name).Count();

                        if (Count_Temp == 0)
                        {
                            Date_Temp = db.WorkCounts.Select(i => i.CreateDate).Distinct().ToList();
                            int Date_Temp_Count = db.WorkCounts.Select(i => i.CreateDate).Distinct().Count();

                            if (Date_Temp_Count == 0){
                                ;
                            }
                            else {
                                Date_Temp_Count = Date_Temp_Count - 1;
                                Date_Temp.RemoveAt(Date_Temp_Count);
                            }
                            
                            foreach (var item_3 in Date_Temp)
                            {
                                obj_temp.Ten_Thiet_Bi = Check;
                                obj_temp.Unit_Name = model.Catelory_Project[j].Unit_Name;
                                obj_temp.Unit_Job = model.Catelory_Project[j].Job;
                                obj_temp.CreateDate = item_3;
                                obj_temp.Unit_Number = 0;
                                db.WorkCounts.Add(obj_temp);
                                db.SaveChanges();
                            }
                        }

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

                    collection.Catelory_Project = model.Catelory_Project;

                    var dates                = new List<DateTime>();
                    var all_number           = new List<int>();
                    var all_job_number_temp  = new List<int>();
                    var all_job_number       = new List<List<int>>();

                    var Date = db.WorkCounts.Where(i => i.CreateDate >= collection.StartDate && i.CreateDate <= collection.EndDate).Select(i => i.CreateDate).Distinct().ToList();
                    var Date_Count = db.WorkCounts.Where(i => i.CreateDate >= collection.StartDate && i.CreateDate <= collection.EndDate).Select(i => i.CreateDate).Distinct().Count();

                    model.List_Job = db.WorkCounts.Where(i => i.CreateDate >= collection.StartDate && i.CreateDate <= collection.EndDate).Select(i => i.Unit_Job).Distinct().ToList();
                    model.Number_Job = db.WorkCounts.Where(i => i.CreateDate >= collection.StartDate && i.CreateDate <= collection.EndDate).Select(i => i.Unit_Job).Distinct().Count();

                    DateTime Date_te = Convert.ToDateTime(Date[0]);
                    var dt = Date_te;

                    List<List<WorkCount>> myList = new List<List<WorkCount>>();

                    for (var temp_test = 0; temp_test < Date_Count; dt = dt.AddDays(1), temp_test++)
                    {
                        DateTime Date_Temp_Test = Convert.ToDateTime(Date[temp_test]);
                        dates.Add(Date_Temp_Test);

                        model.WorkCount = db.WorkCounts.Where(i => i.Ten_Thiet_Bi == collection.SelectedProject.Ten_Thiet_Bi && i.CreateDate == Date_Temp_Test).ToList();
                        myList.Add(model.WorkCount);

                        int temp_all_number = 0;

                        for (var j = 0; j < model.WorkCount.Count(); j++)
                        {
                            temp_all_number = temp_all_number + (int)model.WorkCount[j].Unit_Number;

                        }

                        all_number.Add(temp_all_number);
                        
                        for (var j = 0; j < model.Number_Job; j++)
                        {
                            int temp_job_all_number = 0;
                            int temp_job = 0;
                            
                            for (var i = 0; i < model.WorkCount.Count(); i++)
                            {
                                if (model.WorkCount[i].Unit_Job == model.List_Job[j])
                                {
                                    temp_job_all_number = temp_job_all_number + (int)model.WorkCount[i].Unit_Number;
                                }

                                temp_job = i;
                            }

                            all_job_number_temp.Add(temp_job_all_number);

                        }
                        
                    }

                    List<List<WorkCount>> myList_2 = new List<List<WorkCount>>();

                    foreach(var item in collection.Catelory_Project)
                    {
                        model.WorkCount_Temp_2 = db.WorkCounts.Where(i => i.Ten_Thiet_Bi == collection.SelectedProject.Ten_Thiet_Bi && i.Unit_Name == item.Unit_Name && i.CreateDate>= collection.StartDate && i.CreateDate <= collection.EndDate).OrderBy(i => i.CreateDate).ToList();
                        myList_2.Add(model.WorkCount_Temp_2);
                    }

                    model.WorkCount_List    = myList;
                    model.WorkCount_List_2  = myList_2;
   
                    model.SelectDate = dates;
                    model.Total_number = all_number;

                    model.Total_Job_number_2 = all_job_number_temp;

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
