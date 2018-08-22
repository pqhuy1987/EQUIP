using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Models;
using Models.Framework;

namespace ShopOnline.Controllers
{
    public class Code_EquipController : Controller
    {
        //
        // GET: /CS_tbWorkType/

        public ActionResult Index()
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                Code_EquipViewModel model = new Code_EquipViewModel();
                model.Code_Equip = db.Code_Equip.OrderBy(m => m.ID).ToList();
                model.Code_Group = db.Code_Group.OrderBy(m => m.ID).ToList();

                //--------Add Dropdown for Code_Group-------------------//
                model.Code_Group_All = new List<SelectListItem>();
                var items = new List<SelectListItem>();
                foreach (var Code_Group_Main in model.Code_Group)
                {
                    items.Add(new SelectListItem()
                    {
                        Value = Code_Group_Main.ID.ToString(),
                        Text = Code_Group_Main.Code,
                    });
                }
                model.Code_Group_All = items;
                //--------Add Dropdown for Code_Group-------------------//

                return View(model);
            }
        }

        //
        // GET: /CS_tbWorkType/Details/5

        public ActionResult Details(Code_EquipViewModel collection)
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                if (collection.Code_EquipSelect.ID_Code == 0)
                {
                    Code_EquipViewModel model = new Code_EquipViewModel();
                    model.Code_Equip = db.Code_Equip.OrderBy(i => i.ID).ToList();

                    //--------Add Dropdown for MainProjectName-------------------//
                    model.Code_Group = db.Code_Group.OrderBy(m => m.ID).ToList();
                    model.Code_Group_All = new List<SelectListItem>();
                    var items = new List<SelectListItem>();
                    foreach (var Code_Group in model.Code_Group)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = Code_Group.ID.ToString(),
                            Text = Code_Group.Code,
                        });
                    }
                    model.Code_Group_All = items;
                    //--------Add Dropdown for MainProjectName-------------------//

                    return View("Index", model);
                }
                else
                {
                    Code_EquipViewModel model = new Code_EquipViewModel();
                    model.Code_Equip = db.Code_Equip.Where(m => m.ID_Code == collection.Code_EquipSelect.ID_Code).OrderBy(i => i.ID).ToList();

                    //--------Add Dropdown for MainProjectName-------------------//
                    model.Code_Group = db.Code_Group.OrderBy(m => m.ID).ToList();
                    model.Code_Group_All = new List<SelectListItem>();
                    var items = new List<SelectListItem>();
                    foreach (var Code_Group in model.Code_Group)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = Code_Group.ID.ToString(),
                            Text = Code_Group.Code,
                        });
                    }
                    model.Code_Group_All = items;
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
                Code_EquipViewModel model = new Code_EquipViewModel();
                model.Code_Equip = db.Code_Equip.OrderBy(m => m.ID).ToList();

                //--------Add Dropdown for MainProjectName-------------------//
                model.Code_Group = db.Code_Group.OrderBy(m => m.ID).ToList();
                model.Code_Group_All = new List<SelectListItem>();
                var items = new List<SelectListItem>();
                foreach (var Code_Group in model.Code_Group)
                {
                    items.Add(new SelectListItem()
                    {
                        Value = Code_Group.ID.ToString(),
                        Text = Code_Group.Code,
                    });
                }
                model.Code_Group_All = items;
                //--------Add Dropdown for MainProjectName-------------------//

                return View("Create", model);
            }
        }

        //
        // POST: /CS_tbWorkType/Create

        [HttpPost]
        public ActionResult Create(Code_EquipViewModel collection)
        {
            try
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    Code_Equip obj = new Code_Equip();
                    Code_EquipViewModel model = new Code_EquipViewModel();
                    model.Code_Equip = db.Code_Equip.OrderBy(m => m.ID).ToList();

                    obj.ID_Code = collection.Code_EquipSelect.ID_Code;
                    obj.Equip = collection.Code_EquipSelect.Equip;

                    db.Code_Equip.Add(obj);
                    db.SaveChanges();

                    //--------Add Dropdown for MainProjectName-------------------//
                    model.Code_Group = db.Code_Group.OrderBy(m => m.ID).ToList();
                    model.Code_Group_All = new List<SelectListItem>();
                    var items = new List<SelectListItem>();
                    foreach (var Code_Group in model.Code_Group)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = Code_Group.ID.ToString(),
                            Text = Code_Group.Code,
                        });
                    }
                    model.Code_Group_All = items;
                    //--------Add Dropdown for MainProjectName-------------------//

                    return View("Create", model);
                }
            }
            catch
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    Code_EquipViewModel model = new Code_EquipViewModel();
                    model.Code_Equip = db.Code_Equip.OrderBy(m => m.ID).ToList();

                    //--------Add Dropdown for MainProjectName-------------------//
                    model.Code_Group = db.Code_Group.OrderBy(m => m.ID).ToList();
                    model.Code_Group_All = new List<SelectListItem>();
                    var items = new List<SelectListItem>();
                    foreach (var Code_Group in model.Code_Group)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = Code_Group.ID.ToString(),
                            Text = Code_Group.Code,
                        });
                    }
                    model.Code_Group_All = items;
                    //--------Add Dropdown for MainProjectName-------------------//

                    return View("Create", model);
                }
            }
        }

        //
        // GET: /CS_tbWorkType/Edit/5

        public ActionResult Edit(int id)
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                Code_EquipViewModel model = new Code_EquipViewModel();
                model.Code_EquipSelect = db.Code_Equip.Find(id);

                //--------Add Dropdown for MainProjectName-------------------//
                model.Code_Group = db.Code_Group.OrderBy(m => m.ID).ToList();
                model.Code_Group_All = new List<SelectListItem>();
                var items = new List<SelectListItem>();
                foreach (var Code_Group in model.Code_Group)
                {
                    items.Add(new SelectListItem()
                    {
                        Value = Code_Group.ID.ToString(),
                        Text = Code_Group.Code,
                    });
                }
                model.Code_Group_All = items;
                //--------Add Dropdown for MainProjectName-------------------//

                return View("Edit", model);
            }
        }

        //
        // POST: /CS_tbWorkType/Edit/5

        [HttpPost]
        public ActionResult Save(int id, Code_EquipViewModel collection)
        {
            try
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    Code_EquipViewModel model = new Code_EquipViewModel();
                    Code_Equip Exsiting_Code_Equip = db.Code_Equip.Find(id);
                    Exsiting_Code_Equip.ID_Code = collection.Code_EquipSelect.ID_Code;
                    Exsiting_Code_Equip.Equip = collection.Code_EquipSelect.Equip;
                    db.SaveChanges();

                    //--------Add Dropdown for MainProjectName-------------------//
                    model.Code_Group = db.Code_Group.OrderBy(m => m.ID).ToList();
                    model.Code_Group_All = new List<SelectListItem>();
                    var items = new List<SelectListItem>();
                    foreach (var Code_Group in model.Code_Group)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = Code_Group.ID.ToString(),
                            Text = Code_Group.Code,
                        });
                    }
                    model.Code_Group_All = items;
                    //--------Add Dropdown for MainProjectName-------------------//

                    model.Code_EquipSelect = db.Code_Equip.Find(id);

                    return View("Edit", model);
                }
            }
            catch
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    Code_EquipViewModel model = new Code_EquipViewModel();
                    model.Code_EquipSelect = db.Code_Equip.Find(id);

                    //--------Add Dropdown for MainProjectName-------------------//
                    model.Code_Group = db.Code_Group.OrderBy(m => m.ID).ToList();
                    model.Code_Group_All = new List<SelectListItem>();
                    var items = new List<SelectListItem>();
                    foreach (var Code_Group in model.Code_Group)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = Code_Group.ID.ToString(),
                            Text = Code_Group.Code,
                        });
                    }
                    model.Code_Group_All = items;
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
                Code_EquipViewModel model = new Code_EquipViewModel();
                model.Code_EquipSelect = db.Code_Equip.Find(id);

                //--------Add Dropdown for MainProjectName-------------------//
                model.Code_Group = db.Code_Group.OrderBy(m => m.ID).ToList();
                model.Code_Group_All = new List<SelectListItem>();
                var items = new List<SelectListItem>();
                foreach (var Code_Group in model.Code_Group)
                {
                    items.Add(new SelectListItem()
                    {
                        Value = Code_Group.ID.ToString(),
                        Text = Code_Group.Code,
                    });
                }
                model.Code_Group_All = items;
                //--------Add Dropdown for MainProjectName-------------------//

                return View("Delete", model);
            }
        }

        //
        // POST: /CS_tbWorkType/Delete/5

        [HttpPost]
        public ActionResult Delete(int id, Code_EquipViewModel collection)
        {
            try
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    Code_EquipViewModel model = new Code_EquipViewModel();
                    Code_Equip Exsiting_Code_Equip = db.Code_Equip.Find(id);
                    db.Code_Equip.Remove(Exsiting_Code_Equip);
                    db.SaveChanges();

                    //--------Add Dropdown for MainProjectName-------------------//
                    model.Code_Group = db.Code_Group.OrderBy(m => m.ID).ToList();
                    model.Code_Group_All = new List<SelectListItem>();
                    var items = new List<SelectListItem>();
                    foreach (var Code_Group in model.Code_Group)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = Code_Group.ID.ToString(),
                            Text = Code_Group.Code,
                        });
                    }
                    model.Code_Group_All = items;
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
