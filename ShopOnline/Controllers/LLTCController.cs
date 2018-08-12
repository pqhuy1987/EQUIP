using System;
using System.Data;
using System.Text;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Models;
using Models.Framework;
using Excel = Microsoft.Office.Interop.Excel;


namespace ShopOnline.Controllers
{
    [Authorize]

    public class LLTCController : Controller
    {
        //
        // GET: /LLTC/

        public ActionResult Index()
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                LLTCViewModel model = new LLTCViewModel();
                model.LLTC = db.LLTCs.OrderBy(m => m.ID).ToList();

                //--------Add Dropdown for MainProjectName-------------------//
                model.CS_tbViTri = db.CS_tbViTri.OrderBy(m => m.ID).ToList();
                model.WorkTypeMain_All = new List<SelectListItem>();
                var items = new List<SelectListItem>();
                foreach (var CS_WorkType_Main in model.CS_tbViTri)
                {
                    items.Add(new SelectListItem()
                    {
                        Value = CS_WorkType_Main.ID.ToString(),
                        Text = CS_WorkType_Main.CS_ViTri,
                    });
                }
                model.WorkTypeMain_All = items;
                //--------Add Dropdown for MainProjectName-------------------//
                
                return View(model);
            }
        }

        //
        // GET: /LLTC/Details/5

        public ActionResult DetailsGet(int id)
        {
            //--------Add Dropdown for Project Name-------------------//
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                LLTCViewModel model = new LLTCViewModel();
                //--------Select ID trả kết quả về View-----------//
                model.SelectedLLTC = db.LLTCs.Find(id);
                //--------Add Dropdown for ProjectName-------------------//
                model.LLTC = db.LLTCs.OrderBy(m => m.ID).ToList();
                model.Project = db.Projects.OrderBy(m => m.ID).ToList();
                model.CS_tbLLTCTypeSub = db.CS_tbLLTCTypeSub.Where(m => m.CS_tbLLTC_ID == model.SelectedLLTC.ID).OrderBy(m => m.ID).ToList();
                model.CS_tbWorkType = db.CS_tbWorkType.Where(m => m.CoreWorkType == model.SelectedLLTC.Main_Name_Job).OrderBy(m => m.ID).ToList();
                model.DisplayMode = "Index";
                    //--------Add Dropdown for ProjectName-------------------//
                    model.Project_Name_All = new List<SelectListItem>();
                    var items = new List<SelectListItem>();
                    foreach (var CS_Project_Name in model.Project)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = CS_Project_Name.ID.ToString(),
                            Text = CS_Project_Name.Project_Name,
                        });
                    }
                    model.Project_Name_All = items;
                    //--------Add Dropdown for ProjectName-------------------//

                    //--------Add Dropdown for Details Job-------------------//
                    model.WorkTypeDetails_All = new List<SelectListItem>();
                    var items_2 = new List<SelectListItem>();
                    foreach (var CS_SubJob_Details in model.CS_tbWorkType)
                    {
                        items_2.Add(new SelectListItem()
                        {
                            Value = CS_SubJob_Details.ID.ToString(),
                            Text = CS_SubJob_Details.SubWorkType,
                        });
                    }
                    model.WorkTypeDetails_All = items_2;
                    //--------Add Dropdown for Details Job-------------------//

                return View("Details",model);
            }
            //--------Add Dropdown for Project Name-------------------//

        }
        public ActionResult DetailsEditGet(int id, int LLTC_ID)
        {
            //--------Add Dropdown for Project Name-------------------//
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                LLTCViewModel model = new LLTCViewModel();
                //--------Select ID trả kết quả về View-----------//
                model.CS_tbLLTCTypeSub_Select = db.CS_tbLLTCTypeSub.Find(id);
                model.SelectedLLTC = db.LLTCs.Find(LLTC_ID);
                //--------Add Dropdown for ProjectName-------------------//
                model.LLTC = db.LLTCs.OrderBy(m => m.ID).ToList();
                model.Project = db.Projects.OrderBy(m => m.ID).ToList();
                model.CS_tbLLTCTypeSub = db.CS_tbLLTCTypeSub.Where(m => m.CS_tbLLTC_ID == model.SelectedLLTC.ID).OrderBy(m => m.ID).ToList();
                model.CS_tbWorkType = db.CS_tbWorkType.Where(m => m.CoreWorkType == model.SelectedLLTC.Main_Name_Job).OrderBy(m => m.ID).ToList();
                model.DisplayMode = "Edit";
                //--------Add Dropdown for ProjectName-------------------//
                model.Project_Name_All = new List<SelectListItem>();
                var items = new List<SelectListItem>();
                foreach (var CS_Project_Name in model.Project)
                {
                    items.Add(new SelectListItem()
                    {
                        Value = CS_Project_Name.ID.ToString(),
                        Text = CS_Project_Name.Project_Name,
                    });
                }
                model.Project_Name_All = items;
                //--------Add Dropdown for ProjectName-------------------//

                //--------Add Dropdown for Details Job-------------------//
                model.WorkTypeDetails_All = new List<SelectListItem>();
                var items_2 = new List<SelectListItem>();
                foreach (var CS_SubJob_Details in model.CS_tbWorkType)
                {
                    items_2.Add(new SelectListItem()
                    {
                        Value = CS_SubJob_Details.ID.ToString(),
                        Text = CS_SubJob_Details.SubWorkType,
                    });
                }
                model.WorkTypeDetails_All = items_2;
                //--------Add Dropdown for Details Job-------------------//

                return View("Details", model);
            }
            //--------Add Dropdown for Project Name-------------------//

        }
        public ActionResult DetailsDeleteGet(int id, int LLTC_ID)
        {
            //--------Add Dropdown for Project Name-------------------//
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                LLTCViewModel model = new LLTCViewModel();
                //--------Select ID trả kết quả về View-----------//
                model.CS_tbLLTCTypeSub_Select = db.CS_tbLLTCTypeSub.Find(id);
                model.SelectedLLTC = db.LLTCs.Find(LLTC_ID);
                //--------Add Dropdown for ProjectName-------------------//
                model.LLTC = db.LLTCs.OrderBy(m => m.ID).ToList();
                model.Project = db.Projects.OrderBy(m => m.ID).ToList();
                model.CS_tbLLTCTypeSub = db.CS_tbLLTCTypeSub.Where(m => m.CS_tbLLTC_ID == model.SelectedLLTC.ID).OrderBy(m => m.ID).ToList();
                model.CS_tbWorkType = db.CS_tbWorkType.Where(m => m.CoreWorkType == model.SelectedLLTC.Main_Name_Job).OrderBy(m => m.ID).ToList();
                model.DisplayMode = "Delete";
                //--------Add Dropdown for ProjectName-------------------//
                model.Project_Name_All = new List<SelectListItem>();
                var items = new List<SelectListItem>();
                foreach (var CS_Project_Name in model.Project)
                {
                    items.Add(new SelectListItem()
                    {
                        Value = CS_Project_Name.ID.ToString(),
                        Text = CS_Project_Name.Project_Name,
                    });
                }
                model.Project_Name_All = items;
                //--------Add Dropdown for ProjectName-------------------//

                //--------Add Dropdown for Details Job-------------------//
                model.WorkTypeDetails_All = new List<SelectListItem>();
                var items_2 = new List<SelectListItem>();
                foreach (var CS_SubJob_Details in model.CS_tbWorkType)
                {
                    items_2.Add(new SelectListItem()
                    {
                        Value = CS_SubJob_Details.ID.ToString(),
                        Text = CS_SubJob_Details.SubWorkType,
                    });
                }
                model.WorkTypeDetails_All = items_2;
                //--------Add Dropdown for Details Job-------------------//

                return View("Details", model);
            }
            //--------Add Dropdown for Project Name-------------------//

        }

        [HttpPost]
        public ActionResult DetailsPost(int LLTC_ID, LLTCViewModel collection)
        {
            try
            {
                //--------Add Dropdown for Project Name-------------------//
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    LLTCViewModel model = new LLTCViewModel();
                    CS_tbLLTCTypeSub obj = new CS_tbLLTCTypeSub();

                    obj.CS_tbLLTC_ID = collection.CS_tbLLTCTypeSub_Select.CS_tbLLTC_ID;
                    obj.CS_tbLLTCNameSiteID = collection.CS_tbLLTCTypeSub_Select.CS_tbLLTCNameSiteID;
                    obj.CS_tbLLTCNumberRegisterSub = collection.CS_tbLLTCTypeSub_Select.CS_tbLLTCNumberRegisterSub;
                    obj.CS_tbLLTCNameJobDetailsSub = collection.CS_tbLLTCTypeSub_Select.CS_tbLLTCNameJobDetailsSub;
                    obj.CS_tbLLTCNameSiteManagerSub = collection.CS_tbLLTCTypeSub_Select.CS_tbLLTCNameSiteManagerSub;
                    obj.CS_tbLLTCNameSiteManagerMobileSub = collection.CS_tbLLTCTypeSub_Select.CS_tbLLTCNameSiteManagerMobileSub;
                    obj.CS_tbLLTCStartDateSub = collection.CS_tbLLTCTypeSub_Select.CS_tbLLTCStartDateSub;
                    obj.CS_tbLLTCEndDateSub = collection.CS_tbLLTCTypeSub_Select.CS_tbLLTCEndDateSub;
                    obj.CS_tbLLTCStatusSub = collection.CS_tbLLTCTypeSub_Select.CS_tbLLTCStatusSub;
                    db.CS_tbLLTCTypeSub.Add(obj);
                    db.SaveChanges();

                    //--------Select ID trả kết quả về View-----------//
                    model.SelectedLLTC = db.LLTCs.Find(LLTC_ID);
                    //--------Add Dropdown for ProjectName-------------------//
                    model.LLTC = db.LLTCs.OrderBy(m => m.ID).ToList();
                    model.Project = db.Projects.OrderBy(m => m.ID).ToList();
                    model.CS_tbLLTCTypeSub = db.CS_tbLLTCTypeSub.Where(m => m.CS_tbLLTC_ID == LLTC_ID).OrderBy(m => m.ID).ToList();
                    model.CS_tbWorkType = db.CS_tbWorkType.Where(m => m.CoreWorkType == model.SelectedLLTC.Main_Name_Job).OrderBy(m => m.ID).ToList();
                    model.DisplayMode = "Index";
                    //--------Add Dropdown for ProjectName-------------------//
                    model.Project_Name_All = new List<SelectListItem>();
                    var items = new List<SelectListItem>();
                    foreach (var CS_Project_Name in model.Project)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = CS_Project_Name.ID.ToString(),
                            Text = CS_Project_Name.Project_Name,
                        });
                    }
                    model.Project_Name_All = items;
                    //--------Add Dropdown for ProjectName-------------------//

                    //--------Add Dropdown for Details Job-------------------//
                    model.WorkTypeDetails_All = new List<SelectListItem>();
                    var items_2 = new List<SelectListItem>();
                    foreach (var CS_SubJob_Details in model.CS_tbWorkType)
                    {
                        items_2.Add(new SelectListItem()
                        {
                            Value = CS_SubJob_Details.ID.ToString(),
                            Text = CS_SubJob_Details.SubWorkType,
                        });
                    }
                    model.WorkTypeDetails_All = items_2;
                    //--------Add Dropdown for Details Job-------------------//
                    //--------Add Dropdown for Project Name-------------------//

                    return View("Details", model);
                }
            }
            catch
            {
                //--------Add Dropdown for Project Name-------------------//
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    LLTCViewModel model = new LLTCViewModel();
                    //--------Select ID trả kết quả về View-----------//
                    model.SelectedLLTC = db.LLTCs.Find(LLTC_ID);
                    //--------Add Dropdown for ProjectName-------------------//
                    model.LLTC = db.LLTCs.OrderBy(m => m.ID).ToList();
                    model.Project = db.Projects.OrderBy(m => m.ID).ToList();
                    model.CS_tbLLTCTypeSub = db.CS_tbLLTCTypeSub.Where(m => m.CS_tbLLTC_ID == LLTC_ID).OrderBy(m => m.ID).ToList();
                    model.CS_tbWorkType = db.CS_tbWorkType.Where(m => m.CoreWorkType == model.SelectedLLTC.Main_Name_Job).OrderBy(m => m.ID).ToList();
                    model.DisplayMode = "Index";
                    //--------Add Dropdown for ProjectName-------------------//
                    model.Project_Name_All = new List<SelectListItem>();
                    var items = new List<SelectListItem>();
                    foreach (var CS_Project_Name in model.Project)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = CS_Project_Name.ID.ToString(),
                            Text = CS_Project_Name.Project_Name,
                        });
                    }
                    model.Project_Name_All = items;
                    //--------Add Dropdown for ProjectName-------------------//

                    //--------Add Dropdown for Details Job-------------------//
                    model.WorkTypeDetails_All = new List<SelectListItem>();
                    var items_2 = new List<SelectListItem>();
                    foreach (var CS_SubJob_Details in model.CS_tbWorkType)
                    {
                        items_2.Add(new SelectListItem()
                        {
                            Value = CS_SubJob_Details.ID.ToString(),
                            Text = CS_SubJob_Details.SubWorkType,
                        });
                    }
                    model.WorkTypeDetails_All = items_2;
                    //--------Add Dropdown for Details Job-------------------//
                    //--------Add Dropdown for Project Name-------------------//

                    return View("Details", model);
                }
            }
        }


        [HttpPost]
        public ActionResult DetailsEditPost(int id, int LLTC_ID, LLTCViewModel collection)
        {
            try
            {
                //--------Add Dropdown for Project Name-------------------//
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    LLTCViewModel model = new LLTCViewModel();
                    CS_tbLLTCTypeSub obj = new CS_tbLLTCTypeSub();
                    obj = db.CS_tbLLTCTypeSub.Find(id);
                    
                    obj.CS_tbLLTC_ID = collection.CS_tbLLTCTypeSub_Select.CS_tbLLTC_ID;
                    obj.CS_tbLLTCNameSiteID = collection.CS_tbLLTCTypeSub_Select.CS_tbLLTCNameSiteID;
                    obj.CS_tbLLTCNumberRegisterSub = collection.CS_tbLLTCTypeSub_Select.CS_tbLLTCNumberRegisterSub;
                    obj.CS_tbLLTCNameJobDetailsSub = collection.CS_tbLLTCTypeSub_Select.CS_tbLLTCNameJobDetailsSub;
                    obj.CS_tbLLTCNameSiteManagerSub = collection.CS_tbLLTCTypeSub_Select.CS_tbLLTCNameSiteManagerSub;
                    obj.CS_tbLLTCNameSiteManagerMobileSub = collection.CS_tbLLTCTypeSub_Select.CS_tbLLTCNameSiteManagerMobileSub;
                    obj.CS_tbLLTCStartDateSub = collection.CS_tbLLTCTypeSub_Select.CS_tbLLTCStartDateSub;
                    obj.CS_tbLLTCEndDateSub = collection.CS_tbLLTCTypeSub_Select.CS_tbLLTCEndDateSub;
                    obj.CS_tbLLTCStatusSub = collection.CS_tbLLTCTypeSub_Select.CS_tbLLTCStatusSub;
                    db.SaveChanges();

                    //--------Select ID trả kết quả về View-----------//
                    model.CS_tbLLTCTypeSub_Select = db.CS_tbLLTCTypeSub.Find(id);
                    model.SelectedLLTC = db.LLTCs.Find(LLTC_ID);
                    //--------Add Dropdown for ProjectName-------------------//
                    model.LLTC = db.LLTCs.OrderBy(m => m.ID).ToList();
                    model.Project = db.Projects.OrderBy(m => m.ID).ToList();
                    model.CS_tbLLTCTypeSub = db.CS_tbLLTCTypeSub.Where(m => m.CS_tbLLTC_ID == model.SelectedLLTC.ID).OrderBy(m => m.ID).ToList();
                    model.CS_tbWorkType = db.CS_tbWorkType.Where(m => m.CoreWorkType == model.SelectedLLTC.Main_Name_Job).OrderBy(m => m.ID).ToList();
                    model.DisplayMode = "Edit";
                    //--------Add Dropdown for ProjectName-------------------//
                    model.Project_Name_All = new List<SelectListItem>();
                    var items = new List<SelectListItem>();
                    foreach (var CS_Project_Name in model.Project)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = CS_Project_Name.ID.ToString(),
                            Text = CS_Project_Name.Project_Name,
                        });
                    }
                    model.Project_Name_All = items;
                    //--------Add Dropdown for ProjectName-------------------//

                    //--------Add Dropdown for Details Job-------------------//
                    model.WorkTypeDetails_All = new List<SelectListItem>();
                    var items_2 = new List<SelectListItem>();
                    foreach (var CS_SubJob_Details in model.CS_tbWorkType)
                    {
                        items_2.Add(new SelectListItem()
                        {
                            Value = CS_SubJob_Details.ID.ToString(),
                            Text = CS_SubJob_Details.SubWorkType,
                        });
                    }
                    model.WorkTypeDetails_All = items_2;
                    //--------Add Dropdown for Details Job-------------------//
                    //--------Add Dropdown for Project Name-------------------//

                    return View("Details", model);
                }
            }
            catch
            {
                //--------Add Dropdown for Project Name-------------------//
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    LLTCViewModel model = new LLTCViewModel();
                    //--------Select ID trả kết quả về View-----------//
                    model.CS_tbLLTCTypeSub_Select = db.CS_tbLLTCTypeSub.Find(id);
                    model.SelectedLLTC = db.LLTCs.Find(LLTC_ID);
                    //--------Add Dropdown for ProjectName-------------------//
                    model.LLTC = db.LLTCs.OrderBy(m => m.ID).ToList();
                    model.Project = db.Projects.OrderBy(m => m.ID).ToList();
                    model.CS_tbLLTCTypeSub = db.CS_tbLLTCTypeSub.Where(m => m.CS_tbLLTC_ID == model.SelectedLLTC.ID).OrderBy(m => m.ID).ToList();
                    model.CS_tbWorkType = db.CS_tbWorkType.Where(m => m.CoreWorkType == model.SelectedLLTC.Main_Name_Job).OrderBy(m => m.ID).ToList();
                    model.DisplayMode = "Edit";
                    //--------Add Dropdown for ProjectName-------------------//
                    model.Project_Name_All = new List<SelectListItem>();
                    var items = new List<SelectListItem>();
                    foreach (var CS_Project_Name in model.Project)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = CS_Project_Name.ID.ToString(),
                            Text = CS_Project_Name.Project_Name,
                        });
                    }
                    model.Project_Name_All = items;
                    //--------Add Dropdown for ProjectName-------------------//

                    //--------Add Dropdown for Details Job-------------------//
                    model.WorkTypeDetails_All = new List<SelectListItem>();
                    var items_2 = new List<SelectListItem>();
                    foreach (var CS_SubJob_Details in model.CS_tbWorkType)
                    {
                        items_2.Add(new SelectListItem()
                        {
                            Value = CS_SubJob_Details.ID.ToString(),
                            Text = CS_SubJob_Details.SubWorkType,
                        });
                    }
                    model.WorkTypeDetails_All = items_2;
                    //--------Add Dropdown for Details Job-------------------//
                    //--------Add Dropdown for Project Name-------------------//

                    return View("Details", model);
                }
            }
        }

        [HttpPost]
        public ActionResult DetailsDeletePost(int id, int LLTC_ID, LLTCViewModel collection)
        {
            try
            {
                //--------Add Dropdown for Project Name-------------------//
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    LLTCViewModel model = new LLTCViewModel();

                    CS_tbLLTCTypeSub Exsiting_LLTCTypeSub = db.CS_tbLLTCTypeSub.Find(id);
                    db.CS_tbLLTCTypeSub.Remove(Exsiting_LLTCTypeSub);
                    db.SaveChanges();

                    model.CS_tbLLTCTypeSub_Select = db.CS_tbLLTCTypeSub.Find(id);

                    //--------Select ID trả kết quả về View-----------//
                    //model.CS_tbLLTCTypeSub_Select = db.CS_tbLLTCTypeSub.Find(id);
                    model.SelectedLLTC = db.LLTCs.Find(LLTC_ID);
                    //--------Add Dropdown for ProjectName-------------------//
                    model.LLTC = db.LLTCs.OrderBy(m => m.ID).ToList();
                    model.Project = db.Projects.OrderBy(m => m.ID).ToList();
                    model.CS_tbLLTCTypeSub = db.CS_tbLLTCTypeSub.Where(m => m.CS_tbLLTC_ID == model.SelectedLLTC.ID).OrderBy(m => m.ID).ToList();
                    model.CS_tbWorkType = db.CS_tbWorkType.Where(m => m.CoreWorkType == model.SelectedLLTC.Main_Name_Job).OrderBy(m => m.ID).ToList();
                    model.DisplayMode = "Finish";
                    //--------Add Dropdown for ProjectName-------------------//
                    model.Project_Name_All = new List<SelectListItem>();
                    var items = new List<SelectListItem>();
                    foreach (var CS_Project_Name in model.Project)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = CS_Project_Name.ID.ToString(),
                            Text = CS_Project_Name.Project_Name,
                        });
                    }
                    model.Project_Name_All = items;
                    //--------Add Dropdown for ProjectName-------------------//

                    //--------Add Dropdown for Details Job-------------------//
                    model.WorkTypeDetails_All = new List<SelectListItem>();
                    var items_2 = new List<SelectListItem>();
                    foreach (var CS_SubJob_Details in model.CS_tbWorkType)
                    {
                        items_2.Add(new SelectListItem()
                        {
                            Value = CS_SubJob_Details.ID.ToString(),
                            Text = CS_SubJob_Details.SubWorkType,
                        });
                    }
                    model.WorkTypeDetails_All = items_2;
                    //--------Add Dropdown for Details Job-------------------//
                    //--------Add Dropdown for Project Name-------------------//

                    return View("Details", model);
                }
            }
            catch
            {
                //--------Add Dropdown for Project Name-------------------//
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    LLTCViewModel model = new LLTCViewModel();
                    //--------Select ID trả kết quả về View-----------//
                    //model.CS_tbLLTCTypeSub_Select = db.CS_tbLLTCTypeSub.Find(id);
                    model.SelectedLLTC = db.LLTCs.Find(LLTC_ID);
                    //--------Add Dropdown for ProjectName-------------------//
                    model.LLTC = db.LLTCs.OrderBy(m => m.ID).ToList();
                    model.Project = db.Projects.OrderBy(m => m.ID).ToList();
                    model.CS_tbLLTCTypeSub = db.CS_tbLLTCTypeSub.Where(m => m.CS_tbLLTC_ID == model.SelectedLLTC.ID).OrderBy(m => m.ID).ToList();
                    model.CS_tbWorkType = db.CS_tbWorkType.Where(m => m.CoreWorkType == model.SelectedLLTC.Main_Name_Job).OrderBy(m => m.ID).ToList();
                    model.DisplayMode = "Finish";
                    //--------Add Dropdown for ProjectName-------------------//
                    model.Project_Name_All = new List<SelectListItem>();
                    var items = new List<SelectListItem>();
                    foreach (var CS_Project_Name in model.Project)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = CS_Project_Name.ID.ToString(),
                            Text = CS_Project_Name.Project_Name,
                        });
                    }
                    model.Project_Name_All = items;
                    //--------Add Dropdown for ProjectName-------------------//

                    //--------Add Dropdown for Details Job-------------------//
                    model.WorkTypeDetails_All = new List<SelectListItem>();
                    var items_2 = new List<SelectListItem>();
                    foreach (var CS_SubJob_Details in model.CS_tbWorkType)
                    {
                        items_2.Add(new SelectListItem()
                        {
                            Value = CS_SubJob_Details.ID.ToString(),
                            Text = CS_SubJob_Details.SubWorkType,
                        });
                    }
                    model.WorkTypeDetails_All = items_2;
                    //--------Add Dropdown for Details Job-------------------//
                    //--------Add Dropdown for Project Name-------------------//

                    return View("Details", model);
                }
            }
        }
        //
        // GET: /LLTC/Create

        public ActionResult Create()
        {

            //--------Add Dropdown for Project Name-------------------//
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                LLTCViewModel model = new LLTCViewModel();
                model.LLTC = db.LLTCs.OrderBy(m => m.ID).ToList();
                model.Project = db.Projects.OrderBy(m => m.ID).ToList();
                model.Project_Name_All = new List<SelectListItem>();
                var items = new List<SelectListItem>();

                foreach (var CS_Project_Name in model.Project)
                {
                    items.Add(new SelectListItem()
                    {
                        Value = CS_Project_Name.ID.ToString(),
                        Text = CS_Project_Name.Project_Name,
                    });
                }
                model.Project_Name_All = items;

                //--------Add Dropdown for MainProjectName-------------------//

                //--------Add Dropdown for CoreJob--------------------------//
                model.CS_tbViTri = db.CS_tbViTri.OrderBy(m => m.ID).ToList();
                model.WorkTypeMain_All = new List<SelectListItem>();
                var items_2 = new List<SelectListItem>();
                foreach (var CS_WorkType_Main in model.CS_tbViTri)
                {
                    items_2.Add(new SelectListItem()
                    {
                        Value = CS_WorkType_Main.ID.ToString(),
                        Text = CS_WorkType_Main.CS_ViTri,
                    });
                }
                model.WorkTypeMain_All = items_2;
                //--------Add Dropdown for CoreJob-------------------------//
                return View(model);
            }
            //--------Add Dropdown for Project Name-------------------//
        }

        //
        // POST: /LLTC/Create

        [HttpPost]
        public ActionResult Create(LLTCViewModel collection)
        {
            try
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    LLTC obj = new LLTC();

                    obj.Main_Name_LLTC = collection.SelectedLLTC.Main_Name_LLTC;
                    obj.Main_Name_Ower = collection.SelectedLLTC.Main_Name_Ower;
                    obj.Main_Number = collection.SelectedLLTC.Main_Number;
                    obj.Main_Total_Number = collection.SelectedLLTC.Main_Total_Number;
                    obj.Main_Name_Job = collection.SelectedLLTC.Main_Name_Job;
                    obj.Main_Area       = collection.SelectedLLTC.Main_Area;
                    obj.Main_Status = collection.SelectedLLTC.Main_Status;
                    obj.Main_Rate = collection.SelectedLLTC.Main_Rate;
                    obj.Main_Note       = collection.SelectedLLTC.Main_Note;                                                                            
                    db.LLTCs.Add(obj);
                    db.SaveChanges();

                    //--------Add Dropdown for Project Name-------------------//
                    LLTCViewModel model = new LLTCViewModel();
                    model.LLTC = db.LLTCs.OrderBy(m => m.ID).ToList();
                    model.Project = db.Projects.OrderBy(m => m.ID).ToList();
                    model.Project_Name_All = new List<SelectListItem>();
                    var items = new List<SelectListItem>();

                    foreach (var CS_Project_Name in model.Project)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = CS_Project_Name.ID.ToString(),
                            Text = CS_Project_Name.Project_Name,
                        });
                    }
                    model.Project_Name_All = items;

                    //--------Add Dropdown for CoreJob--------------------------//
                    model.CS_tbViTri = db.CS_tbViTri.OrderBy(m => m.ID).ToList();
                    model.WorkTypeMain_All = new List<SelectListItem>();
                    var items_2 = new List<SelectListItem>();
                    foreach (var CS_WorkType_Main in model.CS_tbViTri)
                    {
                        items_2.Add(new SelectListItem()
                        {
                            Value = CS_WorkType_Main.ID.ToString(),
                            Text = CS_WorkType_Main.CS_ViTri,
                        });
                    }
                    model.WorkTypeMain_All = items_2;
                    //--------Add Dropdown for CoreJob-------------------------//

                    return View(model);
                    //--------Add Dropdown for Project Name-------------------//
                }
            }
            catch
            {
                //--------Add Dropdown for Project Name-------------------//
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    LLTCViewModel model = new LLTCViewModel();
                    model.LLTC = db.LLTCs.OrderBy(m => m.ID).ToList();
                    model.Project = db.Projects.OrderBy(m => m.ID).ToList();
                    model.Project_Name_All = new List<SelectListItem>();
                    var items = new List<SelectListItem>();

                    foreach (var CS_Project_Name in model.Project)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = CS_Project_Name.ID.ToString(),
                            Text = CS_Project_Name.Project_Name,
                        });
                    }
                    model.Project_Name_All = items;

                    //--------Add Dropdown for CoreJob--------------------------//
                    model.CS_tbViTri = db.CS_tbViTri.OrderBy(m => m.ID).ToList();
                    model.WorkTypeMain_All = new List<SelectListItem>();
                    var items_2 = new List<SelectListItem>();
                    foreach (var CS_WorkType_Main in model.CS_tbViTri)
                    {
                        items_2.Add(new SelectListItem()
                        {
                            Value = CS_WorkType_Main.ID.ToString(),
                            Text = CS_WorkType_Main.CS_ViTri,
                        });
                    }
                    model.WorkTypeMain_All = items_2;
                    //--------Add Dropdown for CoreJob-------------------------//

                    return View(model);
                }
                //--------Add Dropdown for Project Name-------------------//
            }
        }

        //
        // GET: /LLTC/Edit/5

        public ActionResult Edit(int id)
        {
            //--------Add Dropdown for Project Name-------------------//
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                LLTCViewModel model = new LLTCViewModel();
                //--------Select ID trả kết quả về View-----------//
                model.SelectedLLTC = db.LLTCs.Find(id);
                //--------Add Dropdown for ProjectName-------------------//
                model.LLTC = db.LLTCs.OrderBy(m => m.ID).ToList();
                model.Project = db.Projects.OrderBy(m => m.ID).ToList();
                model.Project_Name_All = new List<SelectListItem>();
                var items = new List<SelectListItem>();

                foreach (var CS_Project_Name in model.Project)
                {
                    items.Add(new SelectListItem()
                    {
                        Value = CS_Project_Name.ID.ToString(),
                        Text = CS_Project_Name.Project_Name,
                    });
                }
                model.Project_Name_All = items;

                //--------Add Dropdown for CoreJob--------------------------//
                model.CS_tbViTri = db.CS_tbViTri.OrderBy(m => m.ID).ToList();
                model.WorkTypeMain_All = new List<SelectListItem>();
                var items_2 = new List<SelectListItem>();
                foreach (var CS_WorkType_Main in model.CS_tbViTri)
                {
                    items_2.Add(new SelectListItem()
                    {
                        Value = CS_WorkType_Main.ID.ToString(),
                        Text = CS_WorkType_Main.CS_ViTri,
                    });
                }
                model.WorkTypeMain_All = items_2;
                //--------Add Dropdown for CoreJob-------------------------//

                return View(model);
            }
            //--------Add Dropdown for Project Name-------------------//
        }

        //
        // POST: /LLTC/Edit/5

        [HttpPost]
        public ActionResult Save(int id, LLTCViewModel collection)
        {
            try
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    LLTCViewModel model = new LLTCViewModel();

                    LLTC Exsiting_LLTC = db.LLTCs.Find(id);
                    Exsiting_LLTC.Main_Name_LLTC             = collection.SelectedLLTC.Main_Name_LLTC;
                    Exsiting_LLTC.Main_Name_Ower             = collection.SelectedLLTC.Main_Name_Ower;
                    Exsiting_LLTC.Main_Number                = collection.SelectedLLTC.Main_Number;
                    Exsiting_LLTC.Main_Total_Number          = collection.SelectedLLTC.Main_Total_Number;
                    Exsiting_LLTC.Main_Name_Job              = collection.SelectedLLTC.Main_Name_Job;
                    Exsiting_LLTC.Main_Area                  = collection.SelectedLLTC.Main_Area;
                    Exsiting_LLTC.Main_Status                = collection.SelectedLLTC.Main_Status;
                    Exsiting_LLTC.Main_Rate                  = collection.SelectedLLTC.Main_Rate;
                    Exsiting_LLTC.Main_Note                  = collection.SelectedLLTC.Main_Note;
                    db.SaveChanges();

                    //--------Select ID trả kết quả về View-----------//
                    model.SelectedLLTC = db.LLTCs.Find(id);
                    //--------Add Dropdown for ProjectName-------------------//
                    model.LLTC = db.LLTCs.OrderBy(m => m.ID).ToList();
                    model.Project = db.Projects.OrderBy(m => m.ID).ToList();
                    model.Project_Name_All = new List<SelectListItem>();
                    var items = new List<SelectListItem>();

                    foreach (var CS_Project_Name in model.Project)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = CS_Project_Name.ID.ToString(),
                            Text = CS_Project_Name.Project_Name,
                        });
                    }
                    model.Project_Name_All = items;

                    //--------Add Dropdown for CoreJob--------------------------//
                    model.CS_tbViTri = db.CS_tbViTri.OrderBy(m => m.ID).ToList();
                    model.WorkTypeMain_All = new List<SelectListItem>();
                    var items_2 = new List<SelectListItem>();
                    foreach (var CS_WorkType_Main in model.CS_tbViTri)
                    {
                        items_2.Add(new SelectListItem()
                        {
                            Value = CS_WorkType_Main.ID.ToString(),
                            Text = CS_WorkType_Main.CS_ViTri,
                        });
                    }
                    model.WorkTypeMain_All = items_2;
                    //--------Add Dropdown for CoreJob-------------------------//

                    return View("Edit", model);
                }
            }
            catch
            {
                //--------Add Dropdown for Project Name-------------------//
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    LLTCViewModel model = new LLTCViewModel();
                    //--------Select ID trả kết quả về View-----------//
                    model.SelectedLLTC = db.LLTCs.Find(id);
                    //--------Add Dropdown for ProjectName-------------------//
                    model.LLTC = db.LLTCs.OrderBy(m => m.ID).ToList();
                    model.Project = db.Projects.OrderBy(m => m.ID).ToList();
                    model.Project_Name_All = new List<SelectListItem>();
                    var items = new List<SelectListItem>();

                    foreach (var CS_Project_Name in model.Project)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = CS_Project_Name.ID.ToString(),
                            Text = CS_Project_Name.Project_Name,
                        });
                    }
                    model.Project_Name_All = items;

                    //--------Add Dropdown for CoreJob--------------------------//
                    model.CS_tbViTri = db.CS_tbViTri.OrderBy(m => m.ID).ToList();
                    model.WorkTypeMain_All = new List<SelectListItem>();
                    var items_2 = new List<SelectListItem>();
                    foreach (var CS_WorkType_Main in model.CS_tbViTri)
                    {
                        items_2.Add(new SelectListItem()
                        {
                            Value = CS_WorkType_Main.ID.ToString(),
                            Text = CS_WorkType_Main.CS_ViTri,
                        });
                    }
                    model.WorkTypeMain_All = items_2;
                    //--------Add Dropdown for CoreJob-------------------------//

                    return View("Edit",model);
                }
                //--------Add Dropdown for Project Name-------------------//
            }
        }

        //
        // GET: /LLTC/Delete/5

        public ActionResult Delete(int id)
        {
            //--------Add Dropdown for Project Name-------------------//
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                LLTCViewModel model = new LLTCViewModel();
                //--------Select ID trả kết quả về View-----------//
                model.SelectedLLTC = db.LLTCs.Find(id);
                //--------Add Dropdown for ProjectName-------------------//
                model.LLTC = db.LLTCs.OrderBy(m => m.ID).ToList();
                model.Project = db.Projects.OrderBy(m => m.ID).ToList();
                model.Project_Name_All = new List<SelectListItem>();
                var items = new List<SelectListItem>();

                foreach (var CS_Project_Name in model.Project)
                {
                    items.Add(new SelectListItem()
                    {
                        Value = CS_Project_Name.ID.ToString(),
                        Text = CS_Project_Name.Project_Name,
                    });
                }
                model.Project_Name_All = items;

                //--------Add Dropdown for CoreJob--------------------------//
                model.CS_tbViTri = db.CS_tbViTri.OrderBy(m => m.ID).ToList();
                model.WorkTypeMain_All = new List<SelectListItem>();
                var items_2 = new List<SelectListItem>();
                foreach (var CS_WorkType_Main in model.CS_tbViTri)
                {
                    items_2.Add(new SelectListItem()
                    {
                        Value = CS_WorkType_Main.ID.ToString(),
                        Text = CS_WorkType_Main.CS_ViTri,
                    });
                }
                model.WorkTypeMain_All = items_2;
                //--------Add Dropdown for CoreJob-------------------------//

                return View(model);
            }
            //--------Add Dropdown for Project Name-------------------//
        }

        //
        // POST: /LLTC/Delete/5

        [HttpPost]
        public ActionResult Delete(int id, LLTCViewModel collection)
        {
            try
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    LLTCViewModel model = new LLTCViewModel();

                    LLTC Exsiting_LLTC = db.LLTCs.Find(id);
                    db.LLTCs.Remove(Exsiting_LLTC);
                    db.SaveChanges();

                    //--------Add Dropdown for CoreJob--------------------------//
                    model.CS_tbViTri = db.CS_tbViTri.OrderBy(m => m.ID).ToList();
                    model.WorkTypeMain_All = new List<SelectListItem>();
                    var items_2 = new List<SelectListItem>();
                    foreach (var CS_WorkType_Main in model.CS_tbViTri)
                    {
                        items_2.Add(new SelectListItem()
                        {
                            Value = CS_WorkType_Main.ID.ToString(),
                            Text = CS_WorkType_Main.CS_ViTri,
                        });
                    }
                    model.WorkTypeMain_All = items_2;
                    //--------Add Dropdown for CoreJob-------------------------//

                    return View("Finish", model);
                }
            }
            catch
            {
                //--------Add Dropdown for Project Name-------------------//
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    LLTCViewModel model = new LLTCViewModel();
                    //--------Select ID trả kết quả về View-----------//
                    model.SelectedLLTC = db.LLTCs.Find(id);
                    //--------Add Dropdown for ProjectName-------------------//
                    model.LLTC = db.LLTCs.OrderBy(m => m.ID).ToList();
                    model.Project = db.Projects.OrderBy(m => m.ID).ToList();
                    model.Project_Name_All = new List<SelectListItem>();
                    var items = new List<SelectListItem>();

                    foreach (var CS_Project_Name in model.Project)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = CS_Project_Name.ID.ToString(),
                            Text = CS_Project_Name.Project_Name,
                        });
                    }
                    model.Project_Name_All = items;

                    //--------Add Dropdown for CoreJob--------------------------//
                    model.CS_tbViTri = db.CS_tbViTri.OrderBy(m => m.ID).ToList();
                    model.WorkTypeMain_All = new List<SelectListItem>();
                    var items_2 = new List<SelectListItem>();
                    foreach (var CS_WorkType_Main in model.CS_tbViTri)
                    {
                        items_2.Add(new SelectListItem()
                        {
                            Value = CS_WorkType_Main.ID.ToString(),
                            Text = CS_WorkType_Main.CS_ViTri,
                        });
                    }
                    model.WorkTypeMain_All = items_2;
                    //--------Add Dropdown for CoreJob-------------------------//

                    return View(model);
                }
                //--------Add Dropdown for Project Name-------------------//
            }
        }

        public void killExcel()
        {
            System.Diagnostics.Process[] PROC = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            foreach (System.Diagnostics.Process PK in PROC)
            {
                if (PK.MainWindowTitle.Length == 0)
                {
                    PK.Kill();
                }
            }
        }

        [HttpPost]
        public ActionResult Interop()
        {
            //Microsoft.Office.Interop.Excel.Workbook workbook;

            DataTable employeeTable = new DataTable("Employee");
            employeeTable.Columns.Add("Employee ID");
            employeeTable.Columns.Add("Employee Name");
            employeeTable.Rows.Add("1", "ABC");
            employeeTable.Rows.Add("2", "DEF");
            employeeTable.Rows.Add("3", "PQR");
            employeeTable.Rows.Add("4", "XYZ");

            //Create a Department Table
            DataTable departmentTable = new DataTable("Department");
            departmentTable.Columns.Add("Department ID");
            departmentTable.Columns.Add("Department Name");
            departmentTable.Rows.Add("1", "IT");
            departmentTable.Rows.Add("2", "HR");
            departmentTable.Rows.Add("3", "Finance");

            //Create a DataSet with the existing DataTables
            DataSet dataSet = new DataSet("Organization");
            dataSet.Tables.Add(employeeTable);
            dataSet.Tables.Add(departmentTable);

            //Creating Object of Microsoft.Office.Interop.Excel and creating a Workbook
            var excelApp = new Excel.Application();

            //specify the file name where its actually exist  
            string filepath = Server.MapPath(@"~/Reports/Danh_sách_LLTC_theo_công_trường_ba_miền.xlsx");
            string filepathSave = Server.MapPath(@"~/Reports/");


            //excelApp.Visible = true;
            Excel.Workbook WB = excelApp.Workbooks.Open(filepath);

            Excel.Worksheet workSheet = (Excel.Worksheet)excelApp.Worksheets[1]; //creating excel worksheet
            workSheet.Name = "LLTC_Export"; //name of excel file

            //LINQ to get Column of dataset table
            var columnName = dataSet.Tables[0].Columns.Cast<DataColumn>()
                                 .Select(x => x.ColumnName)
                                 .ToArray();
            int i = 2;
            //Adding column name to worksheet
            foreach (var col in columnName)
            {
                i++;
                workSheet.Cells[4, i] = col;
            }

            //Adding records to worksheet
            int j;
            for (i = 4; i < dataSet.Tables[0].Rows.Count; i++)
            {
                for (j = 2; j < dataSet.Tables[0].Columns.Count; j++)
                {
                    workSheet.Cells[i + 2, j + 1] = Convert.ToString(dataSet.Tables[0].Rows[i][j]);
                }
            }

            //Saving the excel file to “e” directory
            excelApp.DisplayAlerts = false;
            workSheet.SaveAs(filepathSave + workSheet.Name);
            WB.Close(0);
            excelApp.Quit();

            try
            {
                string XlsPath = Server.MapPath(@"~/Reports/LLTC_Export.xlsx");
                FileInfo fileDet = new System.IO.FileInfo(XlsPath);
                Response.Clear();
                Response.Charset = "UTF-8";
                Response.ContentEncoding = Encoding.UTF8;
                Response.AddHeader("Content-Disposition", "attachment; filename=" + Server.UrlEncode(fileDet.Name));
                Response.AddHeader("Content-Length", fileDet.Length.ToString());
                Response.ContentType = "application/ms-excel";
                Response.WriteFile(fileDet.FullName);
                Response.End();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            killExcel();
            return RedirectToAction("Index");

        }

    }
}
