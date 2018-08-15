﻿using System;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Models;
using Models.Framework;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace ShopOnline.Controllers
{
    public class ProjectController : Controller
    {
        //
        // GET: /Admin/Project

        public ActionResult Index()
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                //--------Add Dropdown for Type-------------------//
                ProjectViewModel model      = new ProjectViewModel();
                model.Project                       = db.Projects.OrderBy(m => m.ID).ToList();
                model.CS_tbConstructionSiteType     = db.CS_tbConstructionSiteType.OrderBy(m => m.ID).ToList();
                model.Project_Type_All = new List<SelectListItem>();
                var items = new List<SelectListItem>();

                foreach (var CS_tbConstructionSiteType in model.CS_tbConstructionSiteType)
                {
                    items.Add(new SelectListItem()
                    {
                        Value   = CS_tbConstructionSiteType.Type,
                        Text    = CS_tbConstructionSiteType.Type,
                    });
                }

                model.Project_Type_All = items;
                return View(model);
                //--------Add Dropdown for Type-------------------//               
            }
        }

        //
        // GET: /Admin/Project/Details/5

        public ActionResult Details(int id)
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                ProjectViewModel model = new ProjectViewModel();
                //--------Select ID trả kết quả về View-----------//
                model.SelectedProject   = db.Projects.Find(id);
                model.LLTC              = db.LLTCs.OrderBy(m => m.ID).ToList();
                model.CS_tbLLTCTypeSub = db.CS_tbLLTCTypeSub.Where(m => m.CS_tbLLTCNameSiteID == model.SelectedProject.ID).OrderBy(m => new { m.CS_tbLLTCNameJobDetailsSub, m.ID}).ToList();
                model.CS_tbWorkType     = db.CS_tbWorkType.OrderBy(m => m.ID).ToList();
                model.CS_tbViTri = db.CS_tbViTri.OrderBy(m => m.ID).ToList();
                model.Project           = db.Projects.OrderBy(m => m.ID).ToList();

                //--------Add Dropdown for LLTCName-------------------//
                model.LLTC_Name_All = new List<SelectListItem>();
                var items = new List<SelectListItem>();
                foreach (var CS_LLTC_Name in model.LLTC)
                {
                    items.Add(new SelectListItem()
                    {
                        Value = CS_LLTC_Name.ID.ToString(),
                        Text = CS_LLTC_Name.Main_Name_LLTC,
                    });
                }
                model.LLTC_Name_All = items;
                //--------Add Dropdown for LLTCName-------------------//

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

                //--------Add Dropdown for Core Job-------------------//
                model.WorkTypeCore_All = new List<SelectListItem>();
                var items_3 = new List<SelectListItem>();
                foreach (var CS_CoreJob in model.CS_tbViTri)
                {
                    items_3.Add(new SelectListItem()
                    {
                        Value = CS_CoreJob.ID.ToString(),
                        Text = CS_CoreJob.CS_ViTri,
                    });
                }
                model.WorkTypeCore_All = items_3;
                //--------Add Dropdown for Core Job-------------------//

                //--------Add Dropdown for Project All-------------------//
                model.Project_All = new List<SelectListItem>();
                var items_4 = new List<SelectListItem>();
                foreach (var CS_Project in model.Project)
                {
                    items_4.Add(new SelectListItem()
                    {
                        Value = CS_Project.ID.ToString(),
                        Text = CS_Project.Ten_Thiet_Bi ,
                    });
                }
                model.Project_All = items_4;
                //--------Add Dropdown for Project All-------------------//
                model.DisplayMode = "Index";

                return View("Details", model);
            }
        }

        //
        // GET: /Admin/Project/Details/5

        public ActionResult DetailsEditGet(int id, int LLTC_ID)
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                ProjectViewModel model = new ProjectViewModel();
                //--------Select ID trả kết quả về View-----------//
                model.CS_tbLLTCTypeSub_Select = db.CS_tbLLTCTypeSub.Find(LLTC_ID);
                model.SelectedProject = db.Projects.Find(id);
                model.LLTC = db.LLTCs.OrderBy(m => m.ID).ToList();
                model.CS_tbLLTCTypeSub = db.CS_tbLLTCTypeSub.Where(m => m.CS_tbLLTCNameSiteID == model.SelectedProject.ID).OrderBy(m => new { m.CS_tbLLTCNameJobDetailsSub, m.ID }).ToList();
                model.CS_tbWorkType = db.CS_tbWorkType.OrderBy(m => m.ID).ToList();
                model.CS_tbViTri = db.CS_tbViTri.OrderBy(m => m.ID).ToList();
                model.Project = db.Projects.OrderBy(m => m.ID).ToList();

                //--------Add Dropdown for LLTCName-------------------//
                model.LLTC_Name_All = new List<SelectListItem>();
                var items = new List<SelectListItem>();
                foreach (var CS_LLTC_Name in model.LLTC)
                {
                    items.Add(new SelectListItem()
                    {
                        Value = CS_LLTC_Name.ID.ToString(),
                        Text = CS_LLTC_Name.Main_Name_LLTC,
                    });
                }
                model.LLTC_Name_All = items;
                //--------Add Dropdown for LLTCName-------------------//

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

                //--------Add Dropdown for Core Job-------------------//
                model.WorkTypeCore_All = new List<SelectListItem>();
                var items_3 = new List<SelectListItem>();
                foreach (var CS_CoreJob in model.CS_tbViTri)
                {
                    items_3.Add(new SelectListItem()
                    {
                        Value = CS_CoreJob.ID.ToString(),
                        Text = CS_CoreJob.CS_ViTri,
                    });
                }
                model.WorkTypeCore_All = items_3;
                //--------Add Dropdown for Core Job-------------------//

                //--------Add Dropdown for Project All-------------------//
                model.Project_All = new List<SelectListItem>();
                var items_4 = new List<SelectListItem>();
                foreach (var CS_Project in model.Project)
                {
                    items_4.Add(new SelectListItem()
                    {
                        Value = CS_Project.ID.ToString(),
                        Text = CS_Project.Ten_Thiet_Bi,
                    });
                }
                model.Project_All = items_4;
                //--------Add Dropdown for Project All-------------------//
                model.DisplayMode = "Edit";

                return View("Details", model);
            }
        }
        //
        // GET: /Admin/Project/Details/5

        public ActionResult DetailsGetList(int id, ProjectViewModel collection )
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                ProjectViewModel model = new ProjectViewModel();
                //--------Select ID trả kết quả về View-----------//
                model.SelectedProject = db.Projects.Find(id);
                model.LLTC_Select = db.LLTCs.Find(collection.CS_tbLLTCTypeSub_Select.CS_tbLLTC_ID);
                model.CS_tbLLTCTypeSub = db.CS_tbLLTCTypeSub.Where(m => m.CS_tbLLTCNameSiteID == model.SelectedProject.ID).OrderBy(m => new { m.CS_tbLLTCNameJobDetailsSub, m.ID }).ToList();
                model.CS_tbWorkType = db.CS_tbWorkType.Where(m => m.CoreWorkType == model.LLTC_Select.Main_Name_Job).OrderBy(m => m.ID).ToList();
                model.Project = db.Projects.OrderBy(m => m.ID).ToList();
                model.LLTC = db.LLTCs.OrderBy(m => m.ID).ToList();
 
                //--------Add Dropdown for LLTCName-------------------//
                model.LLTC_Name_All = new List<SelectListItem>();
                var items = new List<SelectListItem>();
                foreach (var CS_LLTC_Name in model.LLTC)
                {
                    items.Add(new SelectListItem()
                    {
                        Value = CS_LLTC_Name.ID.ToString(),
                        Text = CS_LLTC_Name.Main_Name_LLTC,
                    });
                }
                model.LLTC_Name_All = items;
                //--------Add Dropdown for LLTCName-------------------//

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

                model.DisplayMode = "Index";

                return View("Details", model);
            }
        }

        //
        // GET: /Admin/Project/Details/5

        public ActionResult DetailsGetEditList(int id, int LLTCSub_ID, ProjectViewModel collection)
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                ProjectViewModel model = new ProjectViewModel();
                //--------Select ID trả kết quả về View-----------//
                model.CS_tbLLTCTypeSub_Select = db.CS_tbLLTCTypeSub.Find(LLTCSub_ID);
                model.SelectedProject = db.Projects.Find(id);
                model.LLTC_Select = db.LLTCs.Find(collection.CS_tbLLTCTypeSub_Select.CS_tbLLTC_ID);
                model.CS_tbLLTCTypeSub = db.CS_tbLLTCTypeSub.Where(m => m.CS_tbLLTCNameSiteID == model.SelectedProject.ID).OrderBy(m => new { m.CS_tbLLTCNameJobDetailsSub, m.ID }).ToList();
                model.CS_tbWorkType = db.CS_tbWorkType.Where(m => m.CoreWorkType == model.LLTC_Select.Main_Name_Job).OrderBy(m => m.ID).ToList();
                model.Project = db.Projects.OrderBy(m => m.ID).ToList();
                model.LLTC = db.LLTCs.OrderBy(m => m.ID).ToList();

                //--------Add Dropdown for LLTCName-------------------//
                model.LLTC_Name_All = new List<SelectListItem>();
                var items = new List<SelectListItem>();
                foreach (var CS_LLTC_Name in model.LLTC)
                {
                    items.Add(new SelectListItem()
                    {
                        Value = CS_LLTC_Name.ID.ToString(),
                        Text = CS_LLTC_Name.Main_Name_LLTC,
                    });
                }
                model.LLTC_Name_All = items;
                //--------Add Dropdown for LLTCName-------------------//

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

                model.DisplayMode = "Edit";

                return View("Details", model);
            }
        }

        [HttpPost]
        public ActionResult DetailsPost(int id, ProjectViewModel collection)
        {
            try
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    ProjectViewModel model = new ProjectViewModel();

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
                    model.SelectedProject = db.Projects.Find(id);
                    model.LLTC_Select = db.LLTCs.Find(collection.CS_tbLLTCTypeSub_Select.CS_tbLLTC_ID);
                    model.CS_tbLLTCTypeSub = db.CS_tbLLTCTypeSub.Where(m => m.CS_tbLLTCNameSiteID == model.SelectedProject.ID).OrderBy(m => new { m.CS_tbLLTCNameJobDetailsSub, m.ID }).ToList();
                    model.CS_tbWorkType = db.CS_tbWorkType.Where(m => m.CoreWorkType == model.LLTC_Select.Main_Name_Job).OrderBy(m => m.ID).ToList();
                    model.Project = db.Projects.OrderBy(m => m.ID).ToList();
                    model.LLTC = db.LLTCs.OrderBy(m => m.ID).ToList();

                    //--------Add Dropdown for LLTCName-------------------//
                    model.LLTC_Name_All = new List<SelectListItem>();
                    var items = new List<SelectListItem>();
                    foreach (var CS_LLTC_Name in model.LLTC)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = CS_LLTC_Name.ID.ToString(),
                            Text = CS_LLTC_Name.Main_Name_LLTC,
                        });
                    }
                    model.LLTC_Name_All = items;
                    //--------Add Dropdown for LLTCName-------------------//

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
                    model.DisplayMode = "Index";

                    return View("Details", model);
                }
            }
            catch
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    ProjectViewModel model = new ProjectViewModel();
                    //--------Select ID trả kết quả về View-----------//
                    model.SelectedProject = db.Projects.Find(id);
                    model.LLTC = db.LLTCs.OrderBy(m => m.ID).ToList();
                    model.CS_tbLLTCTypeSub = db.CS_tbLLTCTypeSub.Where(m => m.CS_tbLLTCNameSiteID == model.SelectedProject.ID).OrderBy(m => new { m.CS_tbLLTCNameJobDetailsSub, m.ID }).ToList();
                    model.CS_tbWorkType = db.CS_tbWorkType.OrderBy(m => m.ID).ToList();
                    model.CS_tbViTri = db.CS_tbViTri.OrderBy(m => m.ID).ToList();
                    model.Project = db.Projects.OrderBy(m => m.ID).ToList();

                    //--------Add Dropdown for LLTCName-------------------//
                    model.LLTC_Name_All = new List<SelectListItem>();
                    var items = new List<SelectListItem>();
                    foreach (var CS_LLTC_Name in model.LLTC)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = CS_LLTC_Name.ID.ToString(),
                            Text = CS_LLTC_Name.Main_Name_LLTC,
                        });
                    }
                    model.LLTC_Name_All = items;
                    //--------Add Dropdown for LLTCName-------------------//

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

                    //--------Add Dropdown for Core Job-------------------//
                    model.WorkTypeCore_All = new List<SelectListItem>();
                    var items_3 = new List<SelectListItem>();
                    foreach (var CS_CoreJob in model.CS_tbViTri)
                    {
                        items_3.Add(new SelectListItem()
                        {
                            Value = CS_CoreJob.ID.ToString(),
                            Text = CS_CoreJob.CS_ViTri,
                        });
                    }
                    model.WorkTypeCore_All = items_3;
                    //--------Add Dropdown for Core Job-------------------//

                    //--------Add Dropdown for Project All-------------------//
                    model.Project_All = new List<SelectListItem>();
                    var items_4 = new List<SelectListItem>();
                    foreach (var CS_Project in model.Project)
                    {
                        items_4.Add(new SelectListItem()
                        {
                            Value = CS_Project.ID.ToString(),
                            Text = CS_Project.Ten_Thiet_Bi,
                        });
                    }
                    model.Project_All = items_4;
                    //--------Add Dropdown for Project All-------------------//
                    model.DisplayMode = "Index";

                    return View("Details", model);
                }
            }
        }

        [HttpPost]
        public ActionResult DetailsEditPost(int id, int LLTCSub_ID, ProjectViewModel collection)
        {
            try
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    ProjectViewModel model = new ProjectViewModel();

                    CS_tbLLTCTypeSub obj = new CS_tbLLTCTypeSub();
                    obj = db.CS_tbLLTCTypeSub.Find(LLTCSub_ID);

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
                    model.CS_tbLLTCTypeSub_Select = db.CS_tbLLTCTypeSub.Find(LLTCSub_ID);
                    model.SelectedProject = db.Projects.Find(id);
                    model.LLTC_Select = db.LLTCs.Find(collection.CS_tbLLTCTypeSub_Select.CS_tbLLTC_ID);
                    model.CS_tbLLTCTypeSub = db.CS_tbLLTCTypeSub.Where(m => m.CS_tbLLTCNameSiteID == model.SelectedProject.ID).OrderBy(m => new { m.CS_tbLLTCNameJobDetailsSub, m.ID }).ToList();
                    model.CS_tbWorkType = db.CS_tbWorkType.Where(m => m.CoreWorkType == model.LLTC_Select.Main_Name_Job).OrderBy(m => m.ID).ToList();
                    model.Project = db.Projects.OrderBy(m => m.ID).ToList();
                    model.LLTC = db.LLTCs.OrderBy(m => m.ID).ToList();

                    //--------Add Dropdown for LLTCName-------------------//
                    model.LLTC_Name_All = new List<SelectListItem>();
                    var items = new List<SelectListItem>();
                    foreach (var CS_LLTC_Name in model.LLTC)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = CS_LLTC_Name.ID.ToString(),
                            Text = CS_LLTC_Name.Main_Name_LLTC,
                        });
                    }
                    model.LLTC_Name_All = items;
                    //--------Add Dropdown for LLTCName-------------------//

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
                    model.DisplayMode = "Edit";

                    return View("Details", model);
                }
            }
            catch
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    ProjectViewModel model = new ProjectViewModel();
                    //--------Select ID trả kết quả về View-----------//
                    model.SelectedProject = db.Projects.Find(id);
                    model.LLTC = db.LLTCs.OrderBy(m => m.ID).ToList();
                    model.CS_tbLLTCTypeSub = db.CS_tbLLTCTypeSub.Where(m => m.CS_tbLLTCNameSiteID == model.SelectedProject.ID).OrderBy(m => new { m.CS_tbLLTCNameJobDetailsSub, m.ID }).ToList();
                    model.CS_tbWorkType = db.CS_tbWorkType.OrderBy(m => m.ID).ToList();
                    model.CS_tbViTri = db.CS_tbViTri.OrderBy(m => m.ID).ToList();
                    model.Project = db.Projects.OrderBy(m => m.ID).ToList();

                    //--------Add Dropdown for LLTCName-------------------//
                    model.LLTC_Name_All = new List<SelectListItem>();
                    var items = new List<SelectListItem>();
                    foreach (var CS_LLTC_Name in model.LLTC)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = CS_LLTC_Name.ID.ToString(),
                            Text = CS_LLTC_Name.Main_Name_LLTC,
                        });
                    }
                    model.LLTC_Name_All = items;
                    //--------Add Dropdown for LLTCName-------------------//

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

                    //--------Add Dropdown for Core Job-------------------//
                    model.WorkTypeCore_All = new List<SelectListItem>();
                    var items_3 = new List<SelectListItem>();
                    foreach (var CS_CoreJob in model.CS_tbViTri)
                    {
                        items_3.Add(new SelectListItem()
                        {
                            Value = CS_CoreJob.ID.ToString(),
                            Text = CS_CoreJob.CS_ViTri,
                        });
                    }
                    model.WorkTypeCore_All = items_3;
                    //--------Add Dropdown for Core Job-------------------//

                    //--------Add Dropdown for Project All-------------------//
                    model.Project_All = new List<SelectListItem>();
                    var items_4 = new List<SelectListItem>();
                    foreach (var CS_Project in model.Project)
                    {
                        items_4.Add(new SelectListItem()
                        {
                            Value = CS_Project.ID.ToString(),
                            Text = CS_Project.Ten_Thiet_Bi,
                        });
                    }
                    model.Project_All = items_4;
                    //--------Add Dropdown for Project All-------------------//
                    model.DisplayMode = "Edit";

                    return View("Details", model);
                }
            }
        }

        [HttpPost]
        public ActionResult DetailsSub(int id, int LLTC_ID, int display)
        {
            //--------Add Dropdown for Project Name-------------------//
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                ProjectViewModel model = new ProjectViewModel();
                //--------Select ID trả kết quả về View-----------//
                if (display != 1)
                {
                    display = 1;
                    model.DisplayModeSub = display;
                }
                else
                {
                    display = 2;
                    model.DisplayModeSub = display;
                }
                model.LLTC_Select = db.LLTCs.Find(LLTC_ID);
                model.SelectedProject = db.Projects.Find(id);
                model.LLTC = db.LLTCs.OrderBy(m => m.ID).ToList();
                model.CS_tbLLTCTypeSub = db.CS_tbLLTCTypeSub.Where(m => m.CS_tbLLTCNameSiteID == model.SelectedProject.ID).OrderBy(m => m.ID).ToList();
                model.CS_tbWorkType = db.CS_tbWorkType.OrderBy(m => m.ID).ToList();

                //--------Add Dropdown for LLTCName-------------------//
                model.LLTC_Name_All = new List<SelectListItem>();
                var items = new List<SelectListItem>();
                foreach (var CS_LLTC_Name in model.LLTC)
                {
                    items.Add(new SelectListItem()
                    {
                        Value = CS_LLTC_Name.ID.ToString(),
                        Text = CS_LLTC_Name.Main_Name_LLTC,
                    });
                }
                model.LLTC_Name_All = items;
                //--------Add Dropdown for LLTCName-------------------//

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

        //
        // GET: /Admin/Project/Create

        public ActionResult Create()
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                //--------Add Dropdown for Type-------------------//
                ProjectViewModel model = new ProjectViewModel();
                model.Project = db.Projects.OrderBy(m => m.ID).ToList();
                model.CS_tbConstructionSiteType = db.CS_tbConstructionSiteType.OrderBy(m => m.ID).ToList();
                model.CS_tbViTri = db.CS_tbViTri.OrderBy(m => m.ID).ToList();
                model.Project_Type_All = new List<SelectListItem>();
                model.Vi_Tri_All = new List<SelectListItem>();
                var items = new List<SelectListItem>();
                var items_2 = new List<SelectListItem>();

                foreach (var CS_tbConstructionSiteType in model.CS_tbConstructionSiteType)
                {
                    items.Add(new SelectListItem()
                    {
                        Value = CS_tbConstructionSiteType.Type,
                        Text = CS_tbConstructionSiteType.Type,
                    });
                }

                model.Project_Type_All = items;

                foreach (var CS_ViTri in model.CS_tbViTri)
                {
                    items_2.Add(new SelectListItem()
                    {
                        Value = CS_ViTri.CS_ViTri,
                        Text = CS_ViTri.CS_ViTri,
                    });
                }

                model.Project_Type_All = items;
                model.Vi_Tri_All = items_2;

                return View(model);
                //--------Add Dropdown for Type-------------------//
            }
        }

        //
        // POST: /Admin/Project/Create

        [HttpPost]
        public ActionResult Create(ProjectViewModel collection, HttpPostedFileBase uploadfile)
        {
            try
            {
                    using (OnlineShopDbContext db = new OnlineShopDbContext())
                    {
                        Project obj             = new Project();
                        obj.Ten_Thiet_Bi        = collection.SelectedProject.Ten_Thiet_Bi;
                        obj.Phong_Ban           = collection.SelectedProject.Phong_Ban;
                        obj.Vi_Tri              = collection.SelectedProject.Vi_Tri;



                        if (uploadfile == null)
                        {
                            string _FileName = "NoImage.jpg";
                            //string _path = Path.Combine(Server.MapPath("~/Assets/images"), _FileName);
                            //uploadfile.SaveAs(_path);
                            obj.Site_Manager = _FileName;
                        }
                        else
                        {
                            string _FileName = Path.GetFileName(uploadfile.FileName);
                            string _path = Path.Combine(Server.MapPath("~/Assets/images"), _FileName);
                            uploadfile.SaveAs(_path);
                            obj.Site_Manager = _FileName;
                        }

                        obj.Site_Address = collection.SelectedProject.Site_Address;
                        db.Projects.Add(obj);
                        db.SaveChanges();

                        //--------Add Dropdown for Type-------------------//
                        ProjectViewModel model = new ProjectViewModel();
                        model.Project = db.Projects.OrderBy(m => m.ID).ToList();
                        model.CS_tbConstructionSiteType = db.CS_tbConstructionSiteType.OrderBy(m => m.ID).ToList();
                        model.CS_tbViTri = db.CS_tbViTri.OrderBy(m => m.ID).ToList();
                        model.Project_Type_All = new List<SelectListItem>();
                        model.Vi_Tri_All = new List<SelectListItem>();
                        var items = new List<SelectListItem>();
                        var items_2 = new List<SelectListItem>();

                        foreach (var CS_tbConstructionSiteType in model.CS_tbConstructionSiteType)
                        {
                            items.Add(new SelectListItem()
                            {
                                Value = CS_tbConstructionSiteType.Type,
                                Text = CS_tbConstructionSiteType.Type,
                            });
                        }
                        model.Project_Type_All = items;

                        foreach (var CS_ViTri in model.CS_tbViTri)
                        {
                            items_2.Add(new SelectListItem()
                            {
                                Value = CS_ViTri.CS_ViTri,
                                Text = CS_ViTri.CS_ViTri,
                            });
                        }
                        model.Vi_Tri_All = items_2;

                        return RedirectToAction("Create", model);
                        //--------Add Dropdown for Type-------------------//
                    }
            }
            catch
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    //--------Add Dropdown for Type-------------------//
                    ProjectViewModel model = new ProjectViewModel();
                    model.Project = db.Projects.OrderBy(m => m.ID).ToList();
                    model.CS_tbConstructionSiteType = db.CS_tbConstructionSiteType.OrderBy(m => m.ID).ToList();
                    model.Project_Type_All = new List<SelectListItem>();
                    model.Vi_Tri_All = new List<SelectListItem>();

                    var items = new List<SelectListItem>();
                    var items_2 = new List<SelectListItem>();


                    foreach (var CS_tbConstructionSiteType in model.CS_tbConstructionSiteType)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = CS_tbConstructionSiteType.Type,
                            Text = CS_tbConstructionSiteType.Type,
                        });
                    }

                    model.Project_Type_All = items;

                    foreach (var CS_ViTri in model.CS_tbViTri)
                    {
                        items_2.Add(new SelectListItem()
                        {
                            Value = CS_ViTri.CS_ViTri,
                            Text = CS_ViTri.CS_ViTri,
                        });
                    }
                    model.Vi_Tri_All = items_2;

                    return View(model);
                    //--------Add Dropdown for Type-------------------//
                }
            }
        }

        //
        // GET: /Admin/Project/Edit/5

        public ActionResult Edit(int id)
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                ProjectViewModel model = new ProjectViewModel();
                    //--------Select ID trả kết quả về View-----------//
                    model.SelectedProject = db.Projects.Find(id);
                    //--------Add Dropdown for Type-------------------//
                //--------Model để phía trên----------------------//
                model.Project = db.Projects.OrderBy(m => m.ID).ToList();
                model.CS_tbConstructionSiteType = db.CS_tbConstructionSiteType.OrderBy(m => m.ID).ToList();
                model.CS_tbViTri = db.CS_tbViTri.OrderBy(m => m.ID).ToList();
                model.Project_Type_All = new List<SelectListItem>();
                model.Vi_Tri_All = new List<SelectListItem>();

                var items = new List<SelectListItem>();
                var items_2 = new List<SelectListItem>();

                foreach (var CS_tbConstructionSiteType in model.CS_tbConstructionSiteType)
                {
                    items.Add(new SelectListItem()
                    {
                        Value = CS_tbConstructionSiteType.Type,
                        Text = CS_tbConstructionSiteType.Type,
                    });
                }
                model.Project_Type_All = items;


                foreach (var CS_ViTri in model.CS_tbViTri)
                {
                    items_2.Add(new SelectListItem()
                    {
                        Value = CS_ViTri.CS_ViTri,
                        Text = CS_ViTri.CS_ViTri,
                    });
                }
                model.Vi_Tri_All = items_2;

                return View("Edit", model);
                //--------Add Dropdown for Type-------------------//
            }
        }

        [HttpPost]
        public ActionResult Save(int id, ProjectViewModel collection, HttpPostedFileBase uploadfile)
        {
            try
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                { 
                    Project Exsiting_Project = db.Projects.Find(id);

                    Exsiting_Project.Ten_Thiet_Bi = collection.SelectedProject.Ten_Thiet_Bi;
                    Exsiting_Project.Phong_Ban = collection.SelectedProject.Phong_Ban;
                    Exsiting_Project.Vi_Tri = collection.SelectedProject.Vi_Tri;

                    if (uploadfile == null)
                    {
                        string _FileName = Exsiting_Project.Site_Manager;
                        //string _path = Path.Combine(Server.MapPath("~/Assets/images"), _FileName);
                        //uploadfile.SaveAs(_path);
                        Exsiting_Project.Site_Manager = _FileName;
                    }
                    else
                    {
                        string _FileName = Path.GetFileName(uploadfile.FileName);
                        string _path = Path.Combine(Server.MapPath("~/Assets/images"), _FileName);
                        uploadfile.SaveAs(_path);
                        Exsiting_Project.Site_Manager = _FileName;
                    }


 
                    db.SaveChanges();

                    //--------Add Dropdown for Type-------------------//
                    ProjectViewModel model = new ProjectViewModel();
                        //--------Select ID trả kết quả về View-----------//
                        model.SelectedProject = db.Projects.Find(id);
                        //--------Select ID trả kết quả về View-----------//
                    model.Project = db.Projects.OrderBy(m => m.ID).ToList();

                    model.CS_tbConstructionSiteType = db.CS_tbConstructionSiteType.OrderBy(m => m.ID).ToList();
                    model.CS_tbViTri = db.CS_tbViTri.OrderBy(m => m.ID).ToList();

                    model.Project_Type_All = new List<SelectListItem>();
                    model.Vi_Tri_All = new List<SelectListItem>();

                    var items = new List<SelectListItem>();
                    var items_2 = new List<SelectListItem>();

                    foreach (var CS_tbConstructionSiteType in model.CS_tbConstructionSiteType)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = CS_tbConstructionSiteType.Type,
                            Text = CS_tbConstructionSiteType.Type,
                        });
                    }

                    model.Project_Type_All = items;

                    foreach (var CS_ViTri in model.CS_tbViTri)
                    {
                        items_2.Add(new SelectListItem()
                        {
                            Value = CS_ViTri.CS_ViTri,
                            Text = CS_ViTri.CS_ViTri,
                        });
                    }
                    model.Vi_Tri_All = items_2;

                    return View("Edit", model);
                    //--------Add Dropdown for Type-------------------//              
                }
            }
            catch
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    //--------Add Dropdown for Type-------------------//
                    ProjectViewModel model = new ProjectViewModel();
                        //--------Select ID trả kết quả về View-----------//
                        model.SelectedProject = db.Projects.Find(id);
                        //--------Select ID trả kết quả về View-----------//
                    model.Project = db.Projects.OrderBy(m => m.ID).ToList();
                    model.CS_tbConstructionSiteType = db.CS_tbConstructionSiteType.OrderBy(m => m.ID).ToList();
                    model.Project_Type_All = new List<SelectListItem>();
                    var items = new List<SelectListItem>();

                    foreach (var CS_tbConstructionSiteType in model.CS_tbConstructionSiteType)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = CS_tbConstructionSiteType.Type,
                            Text = CS_tbConstructionSiteType.Type,
                        });
                    }

                    model.Project_Type_All = items;
                    return View("Edit", model);
                    //--------Add Dropdown for Type-------------------//
                }
            }
        }

        //
        // GET: /Admin/Project/Delete/5

        public ActionResult Delete(int id)
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                //--------Add Dropdown for Type-------------------//
                ProjectViewModel model = new ProjectViewModel();
                    //--------Select ID trả kết quả về View-----------//
                    model.SelectedProject = db.Projects.Find(id);
                    //--------Select ID trả kết quả về View-----------//
                model.Project = db.Projects.OrderBy(m => m.ID).ToList();
                model.CS_tbConstructionSiteType = db.CS_tbConstructionSiteType.OrderBy(m => m.ID).ToList();
                model.Project_Type_All = new List<SelectListItem>();
                var items = new List<SelectListItem>();

                foreach (var CS_tbConstructionSiteType in model.CS_tbConstructionSiteType)
                {
                    items.Add(new SelectListItem()
                    {
                        Value = CS_tbConstructionSiteType.Type,
                        Text = CS_tbConstructionSiteType.Type,
                    });
                }

                model.Project_Type_All = items;
                return View(model);
                //--------Add Dropdown for Type-------------------//
            }
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
                    ProjectViewModel model = new ProjectViewModel();

                    Project Exsiting_Type = db.Projects.Find(id);
                    db.Projects.Remove(Exsiting_Type);
                    db.SaveChanges();

                    return View("Finish", model);
                }
            }
            catch
            {
                using (OnlineShopDbContext db = new OnlineShopDbContext())
                {
                    //--------Add Dropdown for Type-------------------//
                    ProjectViewModel model = new ProjectViewModel();
                        //--------Select ID trả kết quả về View-----------//
                        model.SelectedProject = db.Projects.Find(id);
                        //--------Select ID trả kết quả về View-----------//
                    model.Project = db.Projects.OrderBy(m => m.ID).ToList();
                    model.CS_tbConstructionSiteType = db.CS_tbConstructionSiteType.OrderBy(m => m.ID).ToList();
                    model.Project_Type_All = new List<SelectListItem>();
                    var items = new List<SelectListItem>();

                    foreach (var CS_tbConstructionSiteType in model.CS_tbConstructionSiteType)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = CS_tbConstructionSiteType.Type,
                            Text = CS_tbConstructionSiteType.Type,
                        });
                    }

                    model.Project_Type_All = items;
                    return View(model);
                    //--------Add Dropdown for Type-------------------//
                }
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

        public void Excel_Export_Small_Template()
        {
            List<int> Section_RowNum = new List<int>();

            int current_rownum_right = 5;
            int current_rownum_left = 5;
            int Card_number;
            ProjectViewModel model = new ProjectViewModel();

            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                //--------Add Dropdown for Type-------------------//
                model.Project = db.Projects.OrderBy(m => m.ID).ToList();
                Card_number = db.Projects.OrderBy(m => m.ID).Count();
            }

            //Microsoft.Office.Interop.Excel.Workbook workbook;
            Microsoft.Office.Interop.Excel._Worksheet oSheet;

            var excelApp = new Excel.Application();

            //specify the file name where its actually exist  
            string filepath = Server.MapPath(@"~/Reports/DANH_SACH_QR_CODE.xlsx");
            string filepathSave = Server.MapPath(@"~/Reports/");
            string filepathImageLogo = Server.MapPath(@"~/Assets/files/logo.png");

            Excel.Workbook WB = excelApp.Workbooks.Open(filepath);
            oSheet = (Microsoft.Office.Interop.Excel._Worksheet)WB.ActiveSheet;

            Excel.Worksheet workSheet = (Excel.Worksheet)excelApp.Worksheets[1]; //creating excel worksheet
            workSheet.Name = "QR_Code_Export"; //name of excel file

            for (int i = 0; i < Card_number; i++)
            {
                if (i % 2 == 0)
                {
                    //------------------------------QR_CARD_RIGHT------------------------------//
                    current_rownum_right++;

                    workSheet.get_Range("B" + current_rownum_right, "B" + (current_rownum_right + 2)).Merge();
                    workSheet.get_Range("B" + current_rownum_right, "B" + (current_rownum_right + 2)).BorderAround2();
                    Microsoft.Office.Interop.Excel.Range oRange_logo = (Microsoft.Office.Interop.Excel.Range)workSheet.Cells[current_rownum_right, 2];
                    float Left_logo = (float)((double)oRange_logo.Left);
                    float Top_logo = (float)((double)oRange_logo.Top);
                    const float ImageSize_logo_W = 45;
                    const float ImageSize_logo_H = 24;
                    workSheet.Shapes.AddPicture(filepathImageLogo, MsoTriState.msoFalse, MsoTriState.msoCTrue, Left_logo + 9, Top_logo + 8, ImageSize_logo_W, ImageSize_logo_H);

                    workSheet.get_Range("C" + current_rownum_right, "D" + (current_rownum_right + 2)).Merge();
                    workSheet.get_Range("C" + current_rownum_right, "D" + (current_rownum_right + 2)).BorderAround2();
                    oSheet.Cells[current_rownum_right, 3] = "TEM THIẾT BỊ VĂN PHÒNG";
                    oSheet.Cells[current_rownum_right, 3].Font.Bold = true;

                    current_rownum_right = current_rownum_right + 3;
                    workSheet.get_Range("D" + current_rownum_right, "D" + (current_rownum_right + 3)).Merge();
                    workSheet.get_Range("D" + current_rownum_right, "D" + (current_rownum_right + 3)).BorderAround2();
                    Microsoft.Office.Interop.Excel.Range oRange = (Microsoft.Office.Interop.Excel.Range)workSheet.Cells[current_rownum_right, 4];
                    float Left = (float)((double)oRange.Left);
                    float Top = (float)((double)oRange.Top);
                    const float ImageSize = 48;
                    workSheet.Shapes.AddPicture("https://chart.googleapis.com/chart?chs=150x150&cht=qr&chl=http://cwd.fdcc.com.vn:8888/Project/Edit/" + model.Project[i].ID, MsoTriState.msoFalse, MsoTriState.msoCTrue, Left + 2, Top + 2, ImageSize, ImageSize);

                    foreach (Microsoft.Office.Interop.Excel.Range cell in workSheet.get_Range("B" + current_rownum_right, "C" + (current_rownum_right + 3)))
                    {
                        cell.BorderAround2();
                    }

                    oSheet.Cells[current_rownum_right, 2] = "Tên Thiết bị:";
                    oSheet.Cells[current_rownum_right, 2].Font.Bold = true;
                    oSheet.Cells[current_rownum_right, 3] = model.Project[i].Ten_Thiet_Bi;
                    current_rownum_right++;

                    oSheet.Cells[current_rownum_right, 2] = "Ngày cấp:";
                    oSheet.Cells[current_rownum_right, 2].Font.Bold = true;
                    oSheet.Cells[current_rownum_right, 3] = "13-08-2018";
                    current_rownum_right++;

                    oSheet.Cells[current_rownum_right, 2] = "Phòng/Ban:";
                    oSheet.Cells[current_rownum_right, 2].Font.Bold = true;
                    oSheet.Cells[current_rownum_right, 3] = model.Project[i].Phong_Ban;
                    current_rownum_right++;

                    oSheet.Cells[current_rownum_right, 2] = "Mã Thiết bị:";
                    oSheet.Cells[current_rownum_right, 2].Font.Bold = true;
                    oSheet.Cells[current_rownum_right, 3] = "BAN_0.6x1.2_001";
                    current_rownum_right++;
                    oSheet.Cells[current_rownum_right, 2].RowHeight = 55.55;
                    //------------------------------QR_CARD_RIGHT------------------------------//
                }
                else
                {
                    //------------------------------QR_CARD_LEFT------------------------------//
                    current_rownum_left++;

                    workSheet.get_Range("G" + current_rownum_left, "G" + (current_rownum_left + 2)).Merge();
                    workSheet.get_Range("G" + current_rownum_left, "G" + (current_rownum_left + 2)).BorderAround2();
                    Microsoft.Office.Interop.Excel.Range oRange_logo = (Microsoft.Office.Interop.Excel.Range)workSheet.Cells[current_rownum_left, 7];
                    float Left_logo = (float)((double)oRange_logo.Left);
                    float Top_logo = (float)((double)oRange_logo.Top);
                    const float ImageSize_logo_W = 45;
                    const float ImageSize_logo_H = 24;
                    workSheet.Shapes.AddPicture(filepathImageLogo, MsoTriState.msoFalse, MsoTriState.msoCTrue, Left_logo + 9, Top_logo + 8, ImageSize_logo_W, ImageSize_logo_H);

                    workSheet.get_Range("H" + current_rownum_left, "I" + (current_rownum_left + 2)).Merge();
                    workSheet.get_Range("H" + current_rownum_left, "I" + (current_rownum_left + 2)).BorderAround2();
                    oSheet.Cells[current_rownum_left, 8] = "TEM THIẾT BỊ VĂN PHÒNG";
                    oSheet.Cells[current_rownum_left, 8].Font.Bold = true;

                    current_rownum_left = current_rownum_left + 3;
                    workSheet.get_Range("I" + current_rownum_left, "I" + (current_rownum_left + 3)).Merge();
                    workSheet.get_Range("I" + current_rownum_left, "I" + (current_rownum_left + 3)).BorderAround2();
                    Microsoft.Office.Interop.Excel.Range oRange = (Microsoft.Office.Interop.Excel.Range)workSheet.Cells[current_rownum_left, 9];
                    float Left = (float)((double)oRange.Left);
                    float Top = (float)((double)oRange.Top);
                    const float ImageSize = 48;
                    workSheet.Shapes.AddPicture("https://chart.googleapis.com/chart?chs=150x150&cht=qr&chl=http://cwd.fdcc.com.vn:8888/Project/Edit/" + model.Project[i].ID, MsoTriState.msoFalse, MsoTriState.msoCTrue, Left + 2, Top + 2, ImageSize, ImageSize);

                    foreach (Microsoft.Office.Interop.Excel.Range cell in workSheet.get_Range("G" + current_rownum_left, "H" + (current_rownum_left + 3)))
                    {
                        cell.BorderAround2();
                    }

                    oSheet.Cells[current_rownum_left, 7] = "Tên Thiết bị:";
                    oSheet.Cells[current_rownum_left, 7].Font.Bold = true;
                    oSheet.Cells[current_rownum_left, 8] = model.Project[i].Ten_Thiet_Bi;
                    current_rownum_left++;

                    oSheet.Cells[current_rownum_left, 7] = "Ngày cấp:";
                    oSheet.Cells[current_rownum_left, 7].Font.Bold = true;
                    oSheet.Cells[current_rownum_left, 8] = "13-08-2018";
                    current_rownum_left++;

                    oSheet.Cells[current_rownum_left, 7] = "Phòng/Ban:";
                    oSheet.Cells[current_rownum_left, 7].Font.Bold = true;
                    oSheet.Cells[current_rownum_left, 8] = model.Project[i].Phong_Ban;
                    current_rownum_left++;

                    oSheet.Cells[current_rownum_left, 7] = "Mã Thiết bị:";
                    oSheet.Cells[current_rownum_left, 7].Font.Bold = true;
                    oSheet.Cells[current_rownum_left, 8] = "BAN_0.6x1.2_001";
                    current_rownum_left++;
                    oSheet.Cells[current_rownum_left, 7].RowHeight = 55.55;
                    //------------------------------QR_CARD_LEFT------------------------------//
                }
            }

            //Saving the excel file to “e” directory
            excelApp.DisplayAlerts = false;
            workSheet.SaveAs(filepathSave + workSheet.Name);
            WB.Close(0);
            //excelApp.Visible = true;
            excelApp.Quit();

            try
            {
                string XlsPath = Server.MapPath(@"~/Reports/QR_Code_Export.xlsx");
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
        }

        public void Excel_Export_Large_Template()
        {
            List<int> Section_RowNum = new List<int>();

            int current_rownum_right = 0;
            int current_rownum_mid = 0;
            int current_rownum_left = 0;
            int Card_number;
            ProjectViewModel model = new ProjectViewModel();

            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                //--------Add Dropdown for Type-------------------//
                model.Project = db.Projects.OrderBy(m => m.ID).ToList();
                Card_number = db.Projects.OrderBy(m => m.ID).Count();
            }

            //Microsoft.Office.Interop.Excel.Workbook workbook;
            Microsoft.Office.Interop.Excel._Worksheet oSheet;

            var excelApp = new Excel.Application();

            //specify the file name where its actually exist  
            string filepath = Server.MapPath(@"~/Reports/DANH_SACH_QR_CODE_LARGE.xlsx");
            string filepathSave = Server.MapPath(@"~/Reports/");
            string filepathImageLogo = Server.MapPath(@"~/Assets/files/logo.png");

            Excel.Workbook WB = excelApp.Workbooks.Open(filepath);
            oSheet = (Microsoft.Office.Interop.Excel._Worksheet)WB.ActiveSheet;

            Excel.Worksheet workSheet = (Excel.Worksheet)excelApp.Worksheets[1]; //creating excel worksheet
            workSheet.Name = "QR_Code_Export"; //name of excel file

            for (int i = 0; i < Card_number; i++)
            {
                if (i % 3 == 0)
                {
                    //------------------------------QR_CARD_RIGHT------------------------------//
                    current_rownum_right++;

                    workSheet.get_Range("A" + current_rownum_right, "A" + (current_rownum_right + 2)).Merge();
                    workSheet.get_Range("A" + current_rownum_right, "A" + (current_rownum_right + 2)).BorderAround2();
                    Microsoft.Office.Interop.Excel.Range oRange_logo = (Microsoft.Office.Interop.Excel.Range)workSheet.Cells[current_rownum_right, 1];
                    float Left_logo = (float)((double)oRange_logo.Left);
                    float Top_logo = (float)((double)oRange_logo.Top);
                    const float ImageSize_logo_W = 36;
                    const float ImageSize_logo_H = 18;
                    workSheet.Shapes.AddPicture(filepathImageLogo, MsoTriState.msoFalse, MsoTriState.msoCTrue, Left_logo + 5, Top_logo + 8, ImageSize_logo_W, ImageSize_logo_H);

                    workSheet.get_Range("B" + current_rownum_right, "C" + (current_rownum_right + 2)).Merge();
                    workSheet.get_Range("B" + current_rownum_right, "C" + (current_rownum_right + 2)).BorderAround2();
                    oSheet.Cells[current_rownum_right, 2] = "TEM THIẾT BỊ VĂN PHÒNG";
                    oSheet.Cells[current_rownum_right, 2].Font.Bold = true;

                    current_rownum_right = current_rownum_right + 3;
                    workSheet.get_Range("C" + current_rownum_right, "C" + (current_rownum_right + 4)).Merge();
                    workSheet.get_Range("C" + current_rownum_right, "C" + (current_rownum_right + 4)).BorderAround2();
                    Microsoft.Office.Interop.Excel.Range oRange = (Microsoft.Office.Interop.Excel.Range)workSheet.Cells[current_rownum_right, 3];
                    float Left = (float)((double)oRange.Left);
                    float Top = (float)((double)oRange.Top);
                    const float ImageSize = 36;
                    workSheet.Shapes.AddPicture("https://chart.googleapis.com/chart?chs=150x150&cht=qr&chl=http://cwd.fdcc.com.vn:8888/Project/Edit/" + model.Project[i].ID, MsoTriState.msoFalse, MsoTriState.msoCTrue, Left + 4, Top + 10, ImageSize, ImageSize);

                    foreach (Microsoft.Office.Interop.Excel.Range cell in workSheet.get_Range("A" + current_rownum_right, "B" + (current_rownum_right + 4)))
                    {
                        cell.BorderAround2();
                    }

                    oSheet.Cells[current_rownum_right, 1] = "Tên Thiết bị:";
                    oSheet.Cells[current_rownum_right, 1].Font.Bold = true;
                    oSheet.Cells[current_rownum_right, 2] = model.Project[i].Ten_Thiet_Bi;
                    current_rownum_right++;

                    oSheet.Cells[current_rownum_right, 1] = "Ngày cấp:";
                    oSheet.Cells[current_rownum_right, 1].Font.Bold = true;
                    oSheet.Cells[current_rownum_right, 2] = "13-08-2018";
                    current_rownum_right++;

                    oSheet.Cells[current_rownum_right, 1] = "Phòng/Ban:";
                    oSheet.Cells[current_rownum_right, 1].Font.Bold = true;
                    oSheet.Cells[current_rownum_right, 2] = model.Project[i].Phong_Ban;
                    current_rownum_right++;

                    oSheet.Cells[current_rownum_right, 1] = "Vị Trí:";
                    oSheet.Cells[current_rownum_right, 1].Font.Bold = true;
                    oSheet.Cells[current_rownum_right, 2] = model.Project[i].Vi_Tri;
                    current_rownum_right++;

                    oSheet.Cells[current_rownum_right, 1] = "Mã Thiết bị:";
                    oSheet.Cells[current_rownum_right, 1].Font.Bold = true;
                    oSheet.Cells[current_rownum_right, 2] = "BAN_0.6x1.2_001";
                    current_rownum_right++;

                    oSheet.Cells[current_rownum_right, 1].RowHeight = 24;
                    //------------------------------QR_CARD_RIGHT------------------------------//
                }
                else if (i % 3 == 1)
                {
                    //------------------------------QR_CARD_MIDDLE------------------------------//
                    current_rownum_mid++;

                    workSheet.get_Range("E" + current_rownum_mid, "E" + (current_rownum_mid + 2)).Merge();
                    workSheet.get_Range("E" + current_rownum_mid, "E" + (current_rownum_mid + 2)).BorderAround2();
                    Microsoft.Office.Interop.Excel.Range oRange_logo = (Microsoft.Office.Interop.Excel.Range)workSheet.Cells[current_rownum_mid, 5];
                    float Left_logo = (float)((double)oRange_logo.Left);
                    float Top_logo = (float)((double)oRange_logo.Top);
                    const float ImageSize_logo_W = 36;
                    const float ImageSize_logo_H = 18;
                    workSheet.Shapes.AddPicture(filepathImageLogo, MsoTriState.msoFalse, MsoTriState.msoCTrue, Left_logo + 5, Top_logo + 8, ImageSize_logo_W, ImageSize_logo_H);

                    workSheet.get_Range("F" + current_rownum_mid, "G" + (current_rownum_mid + 2)).Merge();
                    workSheet.get_Range("F" + current_rownum_mid, "G" + (current_rownum_mid + 2)).BorderAround2();
                    oSheet.Cells[current_rownum_mid, 6] = "TEM THIẾT BỊ VĂN PHÒNG";
                    oSheet.Cells[current_rownum_mid, 6].Font.Bold = true;

                    current_rownum_mid = current_rownum_mid + 3;
                    workSheet.get_Range("G" + current_rownum_mid, "G" + (current_rownum_mid + 4)).Merge();
                    workSheet.get_Range("G" + current_rownum_mid, "G" + (current_rownum_mid + 4)).BorderAround2();
                    Microsoft.Office.Interop.Excel.Range oRange = (Microsoft.Office.Interop.Excel.Range)workSheet.Cells[current_rownum_mid, 7];
                    float Left = (float)((double)oRange.Left);
                    float Top = (float)((double)oRange.Top);
                    const float ImageSize = 36;
                    workSheet.Shapes.AddPicture("https://chart.googleapis.com/chart?chs=150x150&cht=qr&chl=http://cwd.fdcc.com.vn:8888/Project/Edit/" + model.Project[i].ID, MsoTriState.msoFalse, MsoTriState.msoCTrue, Left + 4, Top + 10, ImageSize, ImageSize);

                    foreach (Microsoft.Office.Interop.Excel.Range cell in workSheet.get_Range("E" + current_rownum_mid, "F" + (current_rownum_mid + 4)))
                    {
                        cell.BorderAround2();
                    }

                    oSheet.Cells[current_rownum_mid, 5] = "Tên Thiết bị:";
                    oSheet.Cells[current_rownum_mid, 5].Font.Bold = true;
                    oSheet.Cells[current_rownum_mid, 6] = model.Project[i].Ten_Thiet_Bi;
                    current_rownum_mid++;

                    oSheet.Cells[current_rownum_mid, 5] = "Ngày cấp:";
                    oSheet.Cells[current_rownum_mid, 5].Font.Bold = true;
                    oSheet.Cells[current_rownum_mid, 6] = "13-08-2018";
                    current_rownum_mid++;

                    oSheet.Cells[current_rownum_mid, 5] = "Phòng/Ban:";
                    oSheet.Cells[current_rownum_mid, 5].Font.Bold = true;
                    oSheet.Cells[current_rownum_mid, 6] = model.Project[i].Phong_Ban;
                    current_rownum_mid++;

                    oSheet.Cells[current_rownum_mid, 5] = "Vị Trí:";
                    oSheet.Cells[current_rownum_mid, 5].Font.Bold = true;
                    oSheet.Cells[current_rownum_mid, 6] = model.Project[i].Vi_Tri;
                    current_rownum_mid++;

                    oSheet.Cells[current_rownum_mid, 5] = "Mã Thiết bị:";
                    oSheet.Cells[current_rownum_mid, 5].Font.Bold = true;
                    oSheet.Cells[current_rownum_mid, 6] = "BAN_0.6x1.2_001";
                    current_rownum_mid++;
                    oSheet.Cells[current_rownum_mid, 5].RowHeight = 24;
                    //------------------------------QR_CARD_MIDDLE------------------------------//
                }
                else if (i % 3 == 2)
                {
                    //------------------------------QR_CARD_LEFT------------------------------//
                    current_rownum_left++;

                    workSheet.get_Range("I" + current_rownum_left, "I" + (current_rownum_left + 2)).Merge();
                    workSheet.get_Range("I" + current_rownum_left, "I" + (current_rownum_left + 2)).BorderAround2();
                    Microsoft.Office.Interop.Excel.Range oRange_logo = (Microsoft.Office.Interop.Excel.Range)workSheet.Cells[current_rownum_left, 9];
                    float Left_logo = (float)((double)oRange_logo.Left);
                    float Top_logo = (float)((double)oRange_logo.Top);
                    const float ImageSize_logo_W = 36;
                    const float ImageSize_logo_H = 18;
                    workSheet.Shapes.AddPicture(filepathImageLogo, MsoTriState.msoFalse, MsoTriState.msoCTrue, Left_logo + 5, Top_logo + 8, ImageSize_logo_W, ImageSize_logo_H);

                    workSheet.get_Range("J" + current_rownum_left, "K" + (current_rownum_left + 2)).Merge();
                    workSheet.get_Range("J" + current_rownum_left, "K" + (current_rownum_left + 2)).BorderAround2();
                    oSheet.Cells[current_rownum_left, 10] = "TEM THIẾT BỊ VĂN PHÒNG";
                    oSheet.Cells[current_rownum_left, 10].Font.Bold = true;

                    current_rownum_left = current_rownum_left + 3;
                    workSheet.get_Range("K" + current_rownum_left, "K" + (current_rownum_left + 4)).Merge();
                    workSheet.get_Range("K" + current_rownum_left, "K" + (current_rownum_left + 4)).BorderAround2();
                    Microsoft.Office.Interop.Excel.Range oRange = (Microsoft.Office.Interop.Excel.Range)workSheet.Cells[current_rownum_left, 11];
                    float Left = (float)((double)oRange.Left);
                    float Top = (float)((double)oRange.Top);
                    const float ImageSize = 36;
                    workSheet.Shapes.AddPicture("https://chart.googleapis.com/chart?chs=150x150&cht=qr&chl=http://cwd.fdcc.com.vn:8888/Project/Edit/" + model.Project[i].ID, MsoTriState.msoFalse, MsoTriState.msoCTrue, Left + 4, Top + 10, ImageSize, ImageSize);

                    foreach (Microsoft.Office.Interop.Excel.Range cell in workSheet.get_Range("I" + current_rownum_left, "J" + (current_rownum_left + 4)))
                    {
                        cell.BorderAround2();
                    }

                    oSheet.Cells[current_rownum_left, 9] = "Tên Thiết bị:";
                    oSheet.Cells[current_rownum_left, 9].Font.Bold = true;
                    oSheet.Cells[current_rownum_left, 10] = model.Project[i].Ten_Thiet_Bi;
                    current_rownum_left++;

                    oSheet.Cells[current_rownum_left, 9] = "Ngày cấp:";
                    oSheet.Cells[current_rownum_left, 9].Font.Bold = true;
                    oSheet.Cells[current_rownum_left, 10] = "13-08-2018";
                    current_rownum_left++;

                    oSheet.Cells[current_rownum_left, 9] = "Phòng/Ban:";
                    oSheet.Cells[current_rownum_left, 9].Font.Bold = true;
                    oSheet.Cells[current_rownum_left, 10] = model.Project[i].Phong_Ban;
                    current_rownum_left++;

                    oSheet.Cells[current_rownum_left, 9] = "Vị Trí:";
                    oSheet.Cells[current_rownum_left, 9].Font.Bold = true;
                    oSheet.Cells[current_rownum_left, 10] = model.Project[i].Vi_Tri;
                    current_rownum_left++;

                    oSheet.Cells[current_rownum_left, 9] = "Mã Thiết bị:";
                    oSheet.Cells[current_rownum_left, 9].Font.Bold = true;
                    oSheet.Cells[current_rownum_left, 10] = "BAN_0.6x1.2_001";
                    current_rownum_left++;
                    oSheet.Cells[current_rownum_left, 9].RowHeight = 24;
                    //------------------------------QR_CARD_LEFT------------------------------//
                }
            }

            //Saving the excel file to “e” directory
            excelApp.DisplayAlerts = false;
            workSheet.SaveAs(filepathSave + workSheet.Name);
            WB.Close(0);
            //excelApp.Visible = true;
            excelApp.Quit();

            try
            {
                string XlsPath = Server.MapPath(@"~/Reports/QR_Code_Export.xlsx");
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
        }

        [HttpPost]
        public ActionResult Interop()
        {
            //Excel_Export_Small_Template();
            Excel_Export_Large_Template();
            return RedirectToAction("LLTC", "Index");

        }

        System.Data.DataTable Load_LLTC_Excel_Report()
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            SqlConnection conn = null;
            conn = new SqlConnection(string.Format("Data Source=SRBDC.FDC.LOCAL; Initial Catalog=CWD; User id=sa; Password=P@ssw0rd"));
            try
            {
                cmd = new SqlCommand("LLTC_Get_List_By_All_Project", conn);
                cmd.CommandType = CommandType.StoredProcedure;
          
                conn.Open();
                SqlDataReader rd = cmd.ExecuteReader();
                result.Load(rd);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                conn.Close();
                cmd.Dispose();
            }
            return result;
        }

    }
}
