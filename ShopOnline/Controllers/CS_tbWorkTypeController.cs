using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Models;
using Models.Framework;

namespace ShopOnline.Controllers
{
    public class CS_tbWorkTypeController : Controller
    {
        //
        // GET: /CS_tbWorkType/

        public ActionResult Index()
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                CS_tbWorkTypeViewModel model = new CS_tbWorkTypeViewModel();
                model.CS_tbWorkType = db.CS_tbWorkType.OrderBy(m => m.ID).ToList();
                model.CS_tbWorkTypeMain = db.CS_tbWorkTypeMain.OrderBy(m => m.ID).ToList();

                //--------Add Dropdown for MainProjectName-------------------//
                model.WorkTypeMain_All = new List<SelectListItem>();
                var items = new List<SelectListItem>();
                foreach (var CS_WorkType_Main in model.CS_tbWorkTypeMain)
                {
                    items.Add(new SelectListItem()
                    {
                        Value = CS_WorkType_Main.ID.ToString(),
                        Text = CS_WorkType_Main.CS_WorkTypeMain,
                    });
                }
                model.WorkTypeMain_All = items;
                //--------Add Dropdown for MainProjectName-------------------//

                return View(model);
            }
        }

        //
        // GET: /CS_tbWorkType/Details/5

        public ActionResult Details(CS_tbWorkTypeViewModel collection)
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                if (collection.CS_tbWorkTypeSelect.CoreWorkType == 0)
                {
                    CS_tbWorkTypeViewModel model = new CS_tbWorkTypeViewModel();
                    model.CS_tbWorkType = db.CS_tbWorkType.OrderBy(i => i.ID).ToList();

                    //--------Add Dropdown for MainProjectName-------------------//
                    model.CS_tbWorkTypeMain = db.CS_tbWorkTypeMain.OrderBy(m => m.ID).ToList();
                    model.WorkTypeMain_All = new List<SelectListItem>();
                    var items = new List<SelectListItem>();
                    foreach (var CS_WorkType_Main in model.CS_tbWorkTypeMain)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = CS_WorkType_Main.ID.ToString(),
                            Text = CS_WorkType_Main.CS_WorkTypeMain,
                        });
                    }
                    model.WorkTypeMain_All = items;
                    //--------Add Dropdown for MainProjectName-------------------//

                    return View("Index", model);
                }
                else
                {
                    CS_tbWorkTypeViewModel model = new CS_tbWorkTypeViewModel();
                    model.CS_tbWorkType = db.CS_tbWorkType.Where(m => m.CoreWorkType == collection.CS_tbWorkTypeSelect.CoreWorkType).OrderBy(i => i.ID).ToList();

                    //--------Add Dropdown for MainProjectName-------------------//
                    model.CS_tbWorkTypeMain = db.CS_tbWorkTypeMain.OrderBy(m => m.ID).ToList();
                    model.WorkTypeMain_All = new List<SelectListItem>();
                    var items = new List<SelectListItem>();
                    foreach (var CS_WorkType_Main in model.CS_tbWorkTypeMain)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = CS_WorkType_Main.ID.ToString(),
                            Text = CS_WorkType_Main.CS_WorkTypeMain,
                        });
                    }
                    model.WorkTypeMain_All = items;
                    //--------Add Dropdown for MainProjectName-------------------//

                    return View("Index", model);
                }

            }
        }

        //
        // GET: /CS_tbWorkType/Create

        public ActionResult Create()
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                CS_tbWorkTypeViewModel model = new CS_tbWorkTypeViewModel();
                model.CS_tbWorkType = db.CS_tbWorkType.OrderBy(m => m.ID).ToList();

                //--------Add Dropdown for MainProjectName-------------------//
                model.CS_tbWorkTypeMain = db.CS_tbWorkTypeMain.OrderBy(m => m.ID).ToList();
                model.WorkTypeMain_All = new List<SelectListItem>();
                var items = new List<SelectListItem>();
                foreach (var CS_WorkType_Main in model.CS_tbWorkTypeMain)
                {
                    items.Add(new SelectListItem()
                    {
                        Value = CS_WorkType_Main.ID.ToString(),
                        Text = CS_WorkType_Main.CS_WorkTypeMain,
                    });
                }
                model.WorkTypeMain_All = items;
                //--------Add Dropdown for MainProjectName-------------------//

                return View("Create",model);
            }
        }

        //
        // POST: /CS_tbWorkType/Create

        [HttpPost]
        public ActionResult Create(CS_tbWorkTypeViewModel collection)
        {
            try
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    CS_tbWorkType obj = new CS_tbWorkType();
                    CS_tbWorkTypeViewModel model = new CS_tbWorkTypeViewModel();
                    model.CS_tbWorkType = db.CS_tbWorkType.OrderBy(m => m.ID).ToList();

                    obj.CoreWorkType = collection.CS_tbWorkTypeSelect.CoreWorkType;
                    obj.SubWorkType = collection.CS_tbWorkTypeSelect.SubWorkType;

                    db.CS_tbWorkType.Add(obj);
                    db.SaveChanges();

                    //--------Add Dropdown for MainProjectName-------------------//
                    model.CS_tbWorkTypeMain = db.CS_tbWorkTypeMain.OrderBy(m => m.ID).ToList();
                    model.WorkTypeMain_All = new List<SelectListItem>();
                    var items = new List<SelectListItem>();
                    foreach (var CS_WorkType_Main in model.CS_tbWorkTypeMain)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = CS_WorkType_Main.ID.ToString(),
                            Text = CS_WorkType_Main.CS_WorkTypeMain,
                        });
                    }
                    model.WorkTypeMain_All = items;
                    //--------Add Dropdown for MainProjectName-------------------//

                    return View("Create", model);
                }
            }
            catch
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    CS_tbWorkTypeViewModel model = new CS_tbWorkTypeViewModel();
                    model.CS_tbWorkType = db.CS_tbWorkType.OrderBy(m => m.ID).ToList();

                    //--------Add Dropdown for MainProjectName-------------------//
                    model.CS_tbWorkTypeMain = db.CS_tbWorkTypeMain.OrderBy(m => m.ID).ToList();
                    model.WorkTypeMain_All = new List<SelectListItem>();
                    var items = new List<SelectListItem>();
                    foreach (var CS_WorkType_Main in model.CS_tbWorkTypeMain)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = CS_WorkType_Main.ID.ToString(),
                            Text = CS_WorkType_Main.CS_WorkTypeMain,
                        });
                    }
                    model.WorkTypeMain_All = items;
                    //--------Add Dropdown for MainProjectName-------------------//

                    return View("Create",model);
                }
            }
        }

        //
        // GET: /CS_tbWorkType/Edit/5

        public ActionResult Edit(int id)
        {
           using (OnlineShopDbContext db = new OnlineShopDbContext())
           {
                 CS_tbWorkTypeViewModel model = new CS_tbWorkTypeViewModel();
                 model.CS_tbWorkTypeSelect = db.CS_tbWorkType.Find(id);

                 //--------Add Dropdown for MainProjectName-------------------//
                 model.CS_tbWorkTypeMain = db.CS_tbWorkTypeMain.OrderBy(m => m.ID).ToList();
                 model.WorkTypeMain_All = new List<SelectListItem>();
                 var items = new List<SelectListItem>();
                 foreach (var CS_WorkType_Main in model.CS_tbWorkTypeMain)
                 {
                     items.Add(new SelectListItem()
                     {
                         Value = CS_WorkType_Main.ID.ToString(),
                         Text = CS_WorkType_Main.CS_WorkTypeMain,
                     });
                 }
                 model.WorkTypeMain_All = items;
                 //--------Add Dropdown for MainProjectName-------------------//

                 return View("Edit", model);
           }
        }

        //
        // POST: /CS_tbWorkType/Edit/5

        [HttpPost]
        public ActionResult Save(int id, CS_tbWorkTypeViewModel collection)
        {
            try
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    CS_tbWorkTypeViewModel model = new CS_tbWorkTypeViewModel();
                    CS_tbWorkType Exsiting_WorkType = db.CS_tbWorkType.Find(id);
                    Exsiting_WorkType.CoreWorkType = collection.CS_tbWorkTypeSelect.CoreWorkType;
                    Exsiting_WorkType.SubWorkType = collection.CS_tbWorkTypeSelect.SubWorkType;
                    db.SaveChanges();

                    //--------Add Dropdown for MainProjectName-------------------//
                    model.CS_tbWorkTypeMain = db.CS_tbWorkTypeMain.OrderBy(m => m.ID).ToList();
                    model.WorkTypeMain_All = new List<SelectListItem>();
                    var items = new List<SelectListItem>();
                    foreach (var CS_WorkType_Main in model.CS_tbWorkTypeMain)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = CS_WorkType_Main.ID.ToString(),
                            Text = CS_WorkType_Main.CS_WorkTypeMain,
                        });
                    }
                    model.WorkTypeMain_All = items;
                    //--------Add Dropdown for MainProjectName-------------------//

                    model.CS_tbWorkTypeSelect = db.CS_tbWorkType.Find(id);

                    return View("Edit", model);
                }
            }
            catch
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    CS_tbWorkTypeViewModel model = new CS_tbWorkTypeViewModel();
                    model.CS_tbWorkTypeSelect = db.CS_tbWorkType.Find(id);

                    //--------Add Dropdown for MainProjectName-------------------//
                    model.CS_tbWorkTypeMain = db.CS_tbWorkTypeMain.OrderBy(m => m.ID).ToList();
                    model.WorkTypeMain_All = new List<SelectListItem>();
                    var items = new List<SelectListItem>();
                    foreach (var CS_WorkType_Main in model.CS_tbWorkTypeMain)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = CS_WorkType_Main.ID.ToString(),
                            Text = CS_WorkType_Main.CS_WorkTypeMain,
                        });
                    }
                    model.WorkTypeMain_All = items;
                    //--------Add Dropdown for MainProjectName-------------------//

                    return View("Edit", model);
                }
            }
        }

        //
        // GET: /CS_tbWorkType/Delete/5

        public ActionResult Delete(int id)
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                CS_tbWorkTypeViewModel model = new CS_tbWorkTypeViewModel();
                model.CS_tbWorkTypeSelect = db.CS_tbWorkType.Find(id);

                //--------Add Dropdown for MainProjectName-------------------//
                model.CS_tbWorkTypeMain = db.CS_tbWorkTypeMain.OrderBy(m => m.ID).ToList();
                model.WorkTypeMain_All = new List<SelectListItem>();
                var items = new List<SelectListItem>();
                foreach (var CS_WorkType_Main in model.CS_tbWorkTypeMain)
                {
                    items.Add(new SelectListItem()
                    {
                        Value = CS_WorkType_Main.ID.ToString(),
                        Text = CS_WorkType_Main.CS_WorkTypeMain,
                    });
                }
                model.WorkTypeMain_All = items;
                //--------Add Dropdown for MainProjectName-------------------//

                return View("Delete", model);
            }
        }

        //
        // POST: /CS_tbWorkType/Delete/5

        [HttpPost]
        public ActionResult Delete(int id, CS_tbWorkTypeViewModel collection)
        {
            try
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    CS_tbWorkTypeViewModel model = new CS_tbWorkTypeViewModel();
                    CS_tbWorkType Exsiting_WorkType = db.CS_tbWorkType.Find(id);
                    db.CS_tbWorkType.Remove(Exsiting_WorkType);
                    db.SaveChanges();

                    //--------Add Dropdown for MainProjectName-------------------//
                    model.CS_tbWorkTypeMain = db.CS_tbWorkTypeMain.OrderBy(m => m.ID).ToList();
                    model.WorkTypeMain_All = new List<SelectListItem>();
                    var items = new List<SelectListItem>();
                    foreach (var CS_WorkType_Main in model.CS_tbWorkTypeMain)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = CS_WorkType_Main.ID.ToString(),
                            Text = CS_WorkType_Main.CS_WorkTypeMain,
                        });
                    }
                    model.WorkTypeMain_All = items;
                    //--------Add Dropdown for MainProjectName-------------------//

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
