using System;
using System.Diagnostics;
using System.Threading;
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
    public class Thiet_BiController : Controller
    {
        //
        // GET: /Admin/Thiet_Bi

        public ActionResult Index()
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                //--------Add Dropdown for Type-------------------//
                ProjectViewModel model      = new ProjectViewModel();

                model.Select_Phong_Ban      = db.CS_tbPhong_Ban.FirstOrDefault().ID;
                model.Select_Group          = db.Code_Group.FirstOrDefault().ID;

                model.Thiet_Bi_Table        = Load_LLTC_Excel_Report_By_Condition(model.Select_Phong_Ban, model.Select_Group);

                model.CS_tbPhong_Ban        = db.CS_tbPhong_Ban.OrderBy(m => m.ID).ToList();
                model.Phong_Ban_All         = new List<SelectListItem>();
                var items                   = new List<SelectListItem>();

                foreach (var CS_tbPhong_Ban in model.CS_tbPhong_Ban)
                {
                    items.Add(new SelectListItem()
                    {
                        Value   = CS_tbPhong_Ban.ID.ToString(),
                        Text    = CS_tbPhong_Ban.Type,
                    });
                }

                model.Phong_Ban_All = items;

                //--------Add Dropdown for Code_Group-------------------//
                model.Code_Group = db.Code_Group.OrderBy(m => m.ID).ToList();
                model.Code_Group_All = new List<SelectListItem>();
                var items_2 = new List<SelectListItem>();
                foreach (var Code_Group_Main in model.Code_Group)
                {
                    items_2.Add(new SelectListItem()
                    {
                        Value = Code_Group_Main.ID.ToString(),
                        Text = Code_Group_Main.Code,
                    });
                }
                model.Code_Group_All = items_2;
                //--------Add Dropdown for Code_Group-------------------//

                return View(model);
                //--------Add Dropdown for Type-------------------//               
            }
        }

        public ActionResult Index_2()
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                //--------Add Dropdown for Type-------------------//
                ProjectViewModel model      = new ProjectViewModel();

                model.Select_Phong_Ban      = db.CS_tbPhong_Ban.FirstOrDefault().ID;
                model.Select_Group          = db.Code_Group.FirstOrDefault().ID;

                model.Thiet_Bi_Table        = Load_LLTC_Excel_Report_By_Condition(model.Select_Phong_Ban, model.Select_Group);
                
                model.CS_tbPhong_Ban        = db.CS_tbPhong_Ban.OrderBy(m => m.ID).ToList();
                model.Phong_Ban_All         = new List<SelectListItem>();
                var items                   = new List<SelectListItem>();

                foreach (var CS_tbPhong_Ban in model.CS_tbPhong_Ban)
                {
                    items.Add(new SelectListItem()
                    {
                        Value = CS_tbPhong_Ban.ID.ToString(),
                        Text = CS_tbPhong_Ban.Type,
                    });
                }

                model.Phong_Ban_All = items;

                //--------Add Dropdown for Code_Group-------------------//
                model.Code_Group = db.Code_Group.OrderBy(m => m.ID).ToList();
                model.Code_Group_All = new List<SelectListItem>();
                var items_2 = new List<SelectListItem>();
                foreach (var Code_Group_Main in model.Code_Group)
                {
                    items_2.Add(new SelectListItem()
                    {
                        Value = Code_Group_Main.ID.ToString(),
                        Text = Code_Group_Main.Code,
                    });
                }
                model.Code_Group_All = items_2;
                //--------Add Dropdown for Code_Group-------------------//

                return View(model);
                //--------Add Dropdown for Type-------------------//               
            }
        }

        //
        // GET: /Admin/Thiet_Bi/Details/5

        public ActionResult Details(int id)
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                ProjectViewModel model = new ProjectViewModel();
                //--------Select ID trả kết quả về View-----------//
                model.SelectedProject   = db.Thiet_Bis.Find(id);
                model.LLTC              = db.LLTCs.OrderBy(m => m.ID).ToList();
                model.CS_tbLLTCTypeSub = db.CS_tbLLTCTypeSub.Where(m => m.CS_tbLLTCNameSiteID == model.SelectedProject.ID).OrderBy(m => new { m.CS_tbLLTCNameJobDetailsSub, m.ID}).ToList();
                model.CS_tbWorkType     = db.CS_tbWorkType.OrderBy(m => m.ID).ToList();
                model.CS_tbViTri = db.CS_tbViTri.OrderBy(m => m.ID).ToList();
                model.Thiet_Bi           = db.Thiet_Bis.OrderBy(m => m.ID).ToList();

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

                //--------Add Dropdown for Thiet_Bi All-------------------//
                model.Project_All = new List<SelectListItem>();
                var items_4 = new List<SelectListItem>();
                foreach (var CS_Project in model.Thiet_Bi)
                {
                    items_4.Add(new SelectListItem()
                    {
                        Value = CS_Project.ID.ToString(),
                        Text = CS_Project.Ten_Thiet_Bi ,
                    });
                }
                model.Project_All = items_4;
                //--------Add Dropdown for Thiet_Bi All-------------------//
                model.DisplayMode = "Index";

                return View("Details", model);
            }
        }

        //
        // GET: /Admin/Thiet_Bi/Details/5

        public ActionResult DetailsEditGet(int id, int LLTC_ID)
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                ProjectViewModel model = new ProjectViewModel();
                //--------Select ID trả kết quả về View-----------//
                model.CS_tbLLTCTypeSub_Select = db.CS_tbLLTCTypeSub.Find(LLTC_ID);
                model.SelectedProject = db.Thiet_Bis.Find(id);
                model.LLTC = db.LLTCs.OrderBy(m => m.ID).ToList();
                model.CS_tbLLTCTypeSub = db.CS_tbLLTCTypeSub.Where(m => m.CS_tbLLTCNameSiteID == model.SelectedProject.ID).OrderBy(m => new { m.CS_tbLLTCNameJobDetailsSub, m.ID }).ToList();
                model.CS_tbWorkType = db.CS_tbWorkType.OrderBy(m => m.ID).ToList();
                model.CS_tbViTri = db.CS_tbViTri.OrderBy(m => m.ID).ToList();
                model.Thiet_Bi = db.Thiet_Bis.OrderBy(m => m.ID).ToList();

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

                //--------Add Dropdown for Thiet_Bi All-------------------//
                model.Project_All = new List<SelectListItem>();
                var items_4 = new List<SelectListItem>();
                foreach (var CS_Project in model.Thiet_Bi)
                {
                    items_4.Add(new SelectListItem()
                    {
                        Value = CS_Project.ID.ToString(),
                        Text = CS_Project.Ten_Thiet_Bi,
                    });
                }
                model.Project_All = items_4;
                //--------Add Dropdown for Thiet_Bi All-------------------//
                model.DisplayMode = "Edit";

                return View("Details", model);
            }
        }
        //
        // GET: /Admin/Thiet_Bi/Details/5

        public ActionResult DetailsGetList(int id, ProjectViewModel collection )
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                ProjectViewModel model = new ProjectViewModel();
                //--------Select ID trả kết quả về View-----------//
                model.SelectedProject = db.Thiet_Bis.Find(id);
                model.LLTC_Select = db.LLTCs.Find(collection.CS_tbLLTCTypeSub_Select.CS_tbLLTC_ID);
                model.CS_tbLLTCTypeSub = db.CS_tbLLTCTypeSub.Where(m => m.CS_tbLLTCNameSiteID == model.SelectedProject.ID).OrderBy(m => new { m.CS_tbLLTCNameJobDetailsSub, m.ID }).ToList();
                model.CS_tbWorkType = db.CS_tbWorkType.Where(m => m.CoreWorkType == model.LLTC_Select.Main_Name_Job).OrderBy(m => m.ID).ToList();
                model.Thiet_Bi = db.Thiet_Bis.OrderBy(m => m.ID).ToList();
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
        // GET: /Admin/Thiet_Bi/Details/5

        public ActionResult DetailsGetEditList(int id, int LLTCSub_ID, ProjectViewModel collection)
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                ProjectViewModel model = new ProjectViewModel();
                //--------Select ID trả kết quả về View-----------//
                model.CS_tbLLTCTypeSub_Select = db.CS_tbLLTCTypeSub.Find(LLTCSub_ID);
                model.SelectedProject = db.Thiet_Bis.Find(id);
                model.LLTC_Select = db.LLTCs.Find(collection.CS_tbLLTCTypeSub_Select.CS_tbLLTC_ID);
                model.CS_tbLLTCTypeSub = db.CS_tbLLTCTypeSub.Where(m => m.CS_tbLLTCNameSiteID == model.SelectedProject.ID).OrderBy(m => new { m.CS_tbLLTCNameJobDetailsSub, m.ID }).ToList();
                model.CS_tbWorkType = db.CS_tbWorkType.Where(m => m.CoreWorkType == model.LLTC_Select.Main_Name_Job).OrderBy(m => m.ID).ToList();
                model.Thiet_Bi = db.Thiet_Bis.OrderBy(m => m.ID).ToList();
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
                    model.SelectedProject = db.Thiet_Bis.Find(id);
                    model.LLTC_Select = db.LLTCs.Find(collection.CS_tbLLTCTypeSub_Select.CS_tbLLTC_ID);
                    model.CS_tbLLTCTypeSub = db.CS_tbLLTCTypeSub.Where(m => m.CS_tbLLTCNameSiteID == model.SelectedProject.ID).OrderBy(m => new { m.CS_tbLLTCNameJobDetailsSub, m.ID }).ToList();
                    model.CS_tbWorkType = db.CS_tbWorkType.Where(m => m.CoreWorkType == model.LLTC_Select.Main_Name_Job).OrderBy(m => m.ID).ToList();
                    model.Thiet_Bi = db.Thiet_Bis.OrderBy(m => m.ID).ToList();
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
                    model.SelectedProject = db.Thiet_Bis.Find(id);
                    model.LLTC = db.LLTCs.OrderBy(m => m.ID).ToList();
                    model.CS_tbLLTCTypeSub = db.CS_tbLLTCTypeSub.Where(m => m.CS_tbLLTCNameSiteID == model.SelectedProject.ID).OrderBy(m => new { m.CS_tbLLTCNameJobDetailsSub, m.ID }).ToList();
                    model.CS_tbWorkType = db.CS_tbWorkType.OrderBy(m => m.ID).ToList();
                    model.CS_tbViTri = db.CS_tbViTri.OrderBy(m => m.ID).ToList();
                    model.Thiet_Bi = db.Thiet_Bis.OrderBy(m => m.ID).ToList();

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

                    //--------Add Dropdown for Thiet_Bi All-------------------//
                    model.Project_All = new List<SelectListItem>();
                    var items_4 = new List<SelectListItem>();
                    foreach (var CS_Project in model.Thiet_Bi)
                    {
                        items_4.Add(new SelectListItem()
                        {
                            Value = CS_Project.ID.ToString(),
                            Text = CS_Project.Ten_Thiet_Bi,
                        });
                    }
                    model.Project_All = items_4;
                    //--------Add Dropdown for Thiet_Bi All-------------------//
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
                    model.SelectedProject = db.Thiet_Bis.Find(id);
                    model.LLTC_Select = db.LLTCs.Find(collection.CS_tbLLTCTypeSub_Select.CS_tbLLTC_ID);
                    model.CS_tbLLTCTypeSub = db.CS_tbLLTCTypeSub.Where(m => m.CS_tbLLTCNameSiteID == model.SelectedProject.ID).OrderBy(m => new { m.CS_tbLLTCNameJobDetailsSub, m.ID }).ToList();
                    model.CS_tbWorkType = db.CS_tbWorkType.Where(m => m.CoreWorkType == model.LLTC_Select.Main_Name_Job).OrderBy(m => m.ID).ToList();
                    model.Thiet_Bi = db.Thiet_Bis.OrderBy(m => m.ID).ToList();
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
                    model.SelectedProject = db.Thiet_Bis.Find(id);
                    model.LLTC = db.LLTCs.OrderBy(m => m.ID).ToList();
                    model.CS_tbLLTCTypeSub = db.CS_tbLLTCTypeSub.Where(m => m.CS_tbLLTCNameSiteID == model.SelectedProject.ID).OrderBy(m => new { m.CS_tbLLTCNameJobDetailsSub, m.ID }).ToList();
                    model.CS_tbWorkType = db.CS_tbWorkType.OrderBy(m => m.ID).ToList();
                    model.CS_tbViTri = db.CS_tbViTri.OrderBy(m => m.ID).ToList();
                    model.Thiet_Bi = db.Thiet_Bis.OrderBy(m => m.ID).ToList();

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

                    //--------Add Dropdown for Thiet_Bi All-------------------//
                    model.Project_All = new List<SelectListItem>();
                    var items_4 = new List<SelectListItem>();
                    foreach (var CS_Project in model.Thiet_Bi)
                    {
                        items_4.Add(new SelectListItem()
                        {
                            Value = CS_Project.ID.ToString(),
                            Text = CS_Project.Ten_Thiet_Bi,
                        });
                    }
                    model.Project_All = items_4;
                    //--------Add Dropdown for Thiet_Bi All-------------------//
                    model.DisplayMode = "Edit";

                    return View("Details", model);
                }
            }
        }

        [HttpPost]
        public ActionResult DetailsSub(int id, int LLTC_ID, int display)
        {
            //--------Add Dropdown for Thiet_Bi Name-------------------//
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                ProjectViewModel model = new ProjectViewModel();
                //--------Select ID trả kết quả về View-----------//
                if (display != 1)
                {
                    display = 1;
                    //model.DisplayModeSub = display;
                }
                else
                {
                    display = 2;
                    //model.DisplayModeSub = display;
                }
                model.LLTC_Select = db.LLTCs.Find(LLTC_ID);
                model.SelectedProject = db.Thiet_Bis.Find(id);
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
            //--------Add Dropdown for Thiet_Bi Name-------------------//
        }

        //
        // GET: /Admin/Thiet_Bi/Create

        public ActionResult Create()
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                //--------Add Dropdown for Type-------------------//
                ProjectViewModel model = new ProjectViewModel();
                model.Thiet_Bi = db.Thiet_Bis.OrderBy(m => m.ID).ToList();
                model.CS_tbPhong_Ban = db.CS_tbPhong_Ban.OrderBy(m => m.ID).ToList();
                model.CS_tbViTri = db.CS_tbViTri.OrderBy(m => m.ID).ToList();
                model.Phong_Ban_All = new List<SelectListItem>();
                model.Vi_Tri_All = new List<SelectListItem>();
                var items = new List<SelectListItem>();
                var items_2 = new List<SelectListItem>();

                foreach (var CS_tbPhong_Ban in model.CS_tbPhong_Ban)
                {
                    items.Add(new SelectListItem()
                    {
                        Value = CS_tbPhong_Ban.ID.ToString(),
                        Text = CS_tbPhong_Ban.Type,
                    });
                }

                model.Phong_Ban_All = items;

                foreach (var CS_ViTri in model.CS_tbViTri)
                {
                    items_2.Add(new SelectListItem()
                    {
                        Value = CS_ViTri.CS_ViTri,
                        Text = CS_ViTri.CS_ViTri,
                    });
                }

                model.Phong_Ban_All = items;
                model.Vi_Tri_All = items_2;

                return View(model);
                //--------Add Dropdown for Type-------------------//
            }
        }

        //
        // POST: /Admin/Thiet_Bi/Create

        [HttpPost]
        public ActionResult Create(ProjectViewModel collection, HttpPostedFileBase uploadfile)
        {
            try
            {
                    using (OnlineShopDbContext db = new OnlineShopDbContext())
                    {
                        Thiet_Bi obj             = new Thiet_Bi();
                        obj.Ten_Thiet_Bi        = collection.SelectedProject.Ten_Thiet_Bi;
                        obj.Phong_Ban           = collection.SelectedProject.Phong_Ban;
                        obj.Vi_Tri              = collection.SelectedProject.Vi_Tri;

                        if (uploadfile == null)
                        {
                            string _FileName = "NoImage.jpg";
                            //string _path = Path.Combine(Server.MapPath("~/Assets/images"), _FileName);
                            //uploadfile.SaveAs(_path);
                            obj.Hinh_Anh = _FileName;
                        }
                        else
                        {
                            //string _FileName = Path.GetFileName(uploadfile.FileName);
                            string _FileName = string.Concat(Path.GetFileNameWithoutExtension(uploadfile.FileName), DateTime.Now.ToString("_yyyy_MM_dd_HH_mm_ss"), Path.GetExtension(uploadfile.FileName));
                            string _path = Path.Combine(Server.MapPath("~/Assets/images"), _FileName);
                            uploadfile.SaveAs(_path);
                            obj.Hinh_Anh = _FileName;
                        }

                        obj.Ma_Thiet_Bi = collection.SelectedProject.Ma_Thiet_Bi;
                        db.Thiet_Bis.Add(obj);
                        db.SaveChanges();

                        //--------Add Dropdown for Type-------------------//
                        ProjectViewModel model = new ProjectViewModel();
                        model.Thiet_Bi = db.Thiet_Bis.OrderBy(m => m.ID).ToList();
                        model.CS_tbPhong_Ban = db.CS_tbPhong_Ban.OrderBy(m => m.ID).ToList();
                        model.CS_tbViTri = db.CS_tbViTri.OrderBy(m => m.ID).ToList();
                        model.Phong_Ban_All = new List<SelectListItem>();
                        model.Vi_Tri_All = new List<SelectListItem>();
                        var items = new List<SelectListItem>();
                        var items_2 = new List<SelectListItem>();

                        foreach (var CS_tbPhong_Ban in model.CS_tbPhong_Ban)
                        {
                            items.Add(new SelectListItem()
                            {
                                Value = CS_tbPhong_Ban.Type,
                                Text = CS_tbPhong_Ban.Type,
                            });
                        }
                        model.Phong_Ban_All = items;

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
                    model.Thiet_Bi = db.Thiet_Bis.OrderBy(m => m.ID).ToList();
                    model.CS_tbPhong_Ban = db.CS_tbPhong_Ban.OrderBy(m => m.ID).ToList();
                    model.Phong_Ban_All = new List<SelectListItem>();
                    model.Vi_Tri_All = new List<SelectListItem>();

                    var items = new List<SelectListItem>();
                    var items_2 = new List<SelectListItem>();

                    foreach (var CS_tbPhong_Ban in model.CS_tbPhong_Ban)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = CS_tbPhong_Ban.Type,
                            Text = CS_tbPhong_Ban.Type,
                        });
                    }

                    model.Phong_Ban_All = items;

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
        // GET: /Admin/Thiet_Bi/Edit/5

        public ActionResult Edit(int id)
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                ProjectViewModel model = new ProjectViewModel();
                    //--------Select ID trả kết quả về View-----------//
                    model.SelectedProject = db.Thiet_Bis.Find(id);
                    //--------Add Dropdown for Type-------------------//
                //--------Model để phía trên----------------------//
                model.Thiet_Bi = db.Thiet_Bis.OrderBy(m => m.ID).ToList();
                model.CS_tbPhong_Ban = db.CS_tbPhong_Ban.OrderBy(m => m.ID).ToList();
                model.CS_tbViTri = db.CS_tbViTri.OrderBy(m => m.ID).ToList();
                model.Phong_Ban_All = new List<SelectListItem>();
                model.Vi_Tri_All = new List<SelectListItem>();

                var items = new List<SelectListItem>();
                var items_2 = new List<SelectListItem>();

                foreach (var CS_tbPhong_Ban in model.CS_tbPhong_Ban)
                {
                    items.Add(new SelectListItem()
                    {
                        Value = CS_tbPhong_Ban.ID.ToString(),
                        Text = CS_tbPhong_Ban.Type,
                    });
                }
                model.Phong_Ban_All = items;


                foreach (var CS_ViTri in model.CS_tbViTri)
                {
                    items_2.Add(new SelectListItem()
                    {
                        Value = CS_ViTri.ID.ToString(),
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
                    Thiet_Bi Exsiting_Project = db.Thiet_Bis.Find(id);

                    Exsiting_Project.Ten_Thiet_Bi = collection.SelectedProject.Ten_Thiet_Bi;
                    Exsiting_Project.Phong_Ban = collection.SelectedProject.Phong_Ban;
                    Exsiting_Project.Vi_Tri = collection.SelectedProject.Vi_Tri;

                    if (uploadfile == null)
                    {
                        string _FileName = Exsiting_Project.Hinh_Anh;
                        //string _path = Path.Combine(Server.MapPath("~/Assets/images"), _FileName);
                        //uploadfile.SaveAs(_path);
                        Exsiting_Project.Hinh_Anh = _FileName;
                    }
                    else
                    {
                        //string _FileName = Path.GetFileName(uploadfile.FileName);
                        string _FileName = string.Concat(Path.GetFileNameWithoutExtension(uploadfile.FileName), DateTime.Now.ToString("_yyyy_MM_dd_HH_mm_ss"), Path.GetExtension(uploadfile.FileName));
                        string _path = Path.Combine(Server.MapPath("~/Assets/images"), _FileName);
                        uploadfile.SaveAs(_path);
                        Exsiting_Project.Hinh_Anh = _FileName;
                    }

                    db.SaveChanges();

                    //--------Add Dropdown for Type-------------------//
                    ProjectViewModel model = new ProjectViewModel();
                        //--------Select ID trả kết quả về View-----------//
                        model.SelectedProject = db.Thiet_Bis.Find(id);
                        //--------Select ID trả kết quả về View-----------//
                    model.Thiet_Bi = db.Thiet_Bis.OrderBy(m => m.ID).ToList();

                    model.CS_tbPhong_Ban = db.CS_tbPhong_Ban.OrderBy(m => m.ID).ToList();
                    model.CS_tbViTri = db.CS_tbViTri.OrderBy(m => m.ID).ToList();

                    model.Phong_Ban_All = new List<SelectListItem>();
                    model.Vi_Tri_All = new List<SelectListItem>();

                    var items = new List<SelectListItem>();
                    var items_2 = new List<SelectListItem>();

                    foreach (var CS_tbPhong_Ban in model.CS_tbPhong_Ban)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = CS_tbPhong_Ban.ID.ToString(),
                            Text = CS_tbPhong_Ban.Type,
                        });
                    }

                    model.Phong_Ban_All = items;

                    foreach (var CS_ViTri in model.CS_tbViTri)
                    {
                        items_2.Add(new SelectListItem()
                        {
                            Value = CS_ViTri.ID.ToString(),
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
                        model.SelectedProject = db.Thiet_Bis.Find(id);
                        //--------Select ID trả kết quả về View-----------//
                    model.Thiet_Bi = db.Thiet_Bis.OrderBy(m => m.ID).ToList();
                    model.CS_tbPhong_Ban = db.CS_tbPhong_Ban.OrderBy(m => m.ID).ToList();
                    model.Phong_Ban_All = new List<SelectListItem>();
                    var items = new List<SelectListItem>();

                    foreach (var CS_tbPhong_Ban in model.CS_tbPhong_Ban)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = CS_tbPhong_Ban.Type,
                            Text = CS_tbPhong_Ban.Type,
                        });
                    }

                    model.Phong_Ban_All = items;
                    return View("Edit", model);
                    //--------Add Dropdown for Type-------------------//
                }
            }
        }

        //
        // GET: /Admin/Thiet_Bi/Delete/5

        public ActionResult Delete(int id)
        {
            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                //--------Add Dropdown for Type-------------------//
                ProjectViewModel model = new ProjectViewModel();
                    //--------Select ID trả kết quả về View-----------//
                    model.SelectedProject = db.Thiet_Bis.Find(id);
                    //--------Select ID trả kết quả về View-----------//
                model.Thiet_Bi = db.Thiet_Bis.OrderBy(m => m.ID).ToList();
                model.CS_tbPhong_Ban = db.CS_tbPhong_Ban.OrderBy(m => m.ID).ToList();
                model.Phong_Ban_All = new List<SelectListItem>();
                var items = new List<SelectListItem>();

                foreach (var CS_tbPhong_Ban in model.CS_tbPhong_Ban)
                {
                    items.Add(new SelectListItem()
                    {
                        Value = CS_tbPhong_Ban.Type,
                        Text = CS_tbPhong_Ban.Type,
                    });
                }

                model.Phong_Ban_All = items;
                return View(model);
                //--------Add Dropdown for Type-------------------//
            }
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
                    ProjectViewModel model = new ProjectViewModel();

                    Thiet_Bi Exsiting_Type = db.Thiet_Bis.Find(id);
                    db.Thiet_Bis.Remove(Exsiting_Type);
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
                        model.SelectedProject = db.Thiet_Bis.Find(id);
                        //--------Select ID trả kết quả về View-----------//
                    model.Thiet_Bi = db.Thiet_Bis.OrderBy(m => m.ID).ToList();
                    model.CS_tbPhong_Ban = db.CS_tbPhong_Ban.OrderBy(m => m.ID).ToList();
                    model.Phong_Ban_All = new List<SelectListItem>();
                    var items = new List<SelectListItem>();

                    foreach (var CS_tbPhong_Ban in model.CS_tbPhong_Ban)
                    {
                        items.Add(new SelectListItem()
                        {
                            Value = CS_tbPhong_Ban.Type,
                            Text = CS_tbPhong_Ban.Type,
                        });
                    }

                    model.Phong_Ban_All = items;
                    return View(model);
                    //--------Add Dropdown for Type-------------------//
                }
            }
        }

        public void killExcel()
        {
            Process[] excelProcsOld = Process.GetProcessesByName("EXCEL");
            Excel.Application myExcelApp = null;
            Excel.Workbooks excelWorkbookTemplate = null;
            Excel.Workbook excelWorkbook = null;
            try
            {
                //DO sth using myExcelApp , excelWorkbookTemplate, excelWorkbook
            }
            catch (Exception ex)
            {
            }
            finally
            {
                //Compare the EXCEL ID and Kill it 
                Process[] excelProcsNew = Process.GetProcessesByName("EXCEL");
                foreach (Process procNew in excelProcsNew)
                {
                    int exist = 0;
                    foreach (Process procOld in excelProcsOld)
                    {
                        if (procNew.Id == procOld.Id)
                        {
                            exist++;
                        }
                    }
                    if (exist == 0)
                    {
                        procNew.Kill();
                    }
                }
            }
        }

        public void Excel_Export_Small_Template(int Phong_Ban, int Group_Code)
        {
            List<int> Section_RowNum = new List<int>();

            int current_rownum_1 = 3;
            int current_rownum_2 = 3;
            int current_rownum_3 = 3;
            int current_rownum_4 = 3;
            int Card_number;
            ProjectViewModel model = new ProjectViewModel();

            model.Thiet_Bi_Table = Load_LLTC_Excel_Report_By_Condition(Phong_Ban, Group_Code);
            Card_number = model.Thiet_Bi_Table.Rows.Count;

            DataRow[] DataRow = model.Thiet_Bi_Table.Select();

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

            var xlCells = workSheet.Cells;
            Excel.Range EntireRow = xlCells.EntireRow;
            EntireRow.RowHeight = 5;

            current_rownum_1++;
            oSheet.Cells[current_rownum_1, 2].RowHeight = 2;
            current_rownum_2++;
            oSheet.Cells[current_rownum_2, 6].RowHeight = 2;
            current_rownum_3++;
            oSheet.Cells[current_rownum_3, 10].RowHeight = 2;
            current_rownum_4++;
            oSheet.Cells[current_rownum_4, 14].RowHeight = 2;

            for (int i = 0; i < Card_number; i++)
            {
                if (i % 4 == 0)
                {
                    //------------------------------QR_CARD_1------------------------------//
                    current_rownum_1++;

                    workSheet.get_Range("B" + current_rownum_1, "B" + (current_rownum_1 + 2)).Merge();
                    workSheet.get_Range("B" + current_rownum_1, "B" + (current_rownum_1 + 2)).BorderAround2();
                    Microsoft.Office.Interop.Excel.Range oRange_logo = (Microsoft.Office.Interop.Excel.Range)workSheet.Cells[current_rownum_1, 2];
                    float Left_logo = (float)((double)oRange_logo.Left);
                    float Top_logo = (float)((double)oRange_logo.Top);
                    const float ImageSize_logo_W = 16;
                    const float ImageSize_logo_H = 9;
                    workSheet.Shapes.AddPicture(filepathImageLogo, MsoTriState.msoFalse, MsoTriState.msoCTrue, Left_logo + 4, Top_logo + 3, ImageSize_logo_W, ImageSize_logo_H);

                    workSheet.get_Range("C" + current_rownum_1, "C" + (current_rownum_1 + 2)).Merge();
                    workSheet.get_Range("C" + current_rownum_1, "C" + (current_rownum_1 + 2)).BorderAround2();
                    oSheet.Cells[current_rownum_1, 3] = "TEM THIẾT BỊ VĂN PHÒNG";
                    oSheet.Cells[current_rownum_1, 3].Font.Bold = true;

                    workSheet.get_Range("D" + current_rownum_1, "D" + (current_rownum_1 + 7)).Merge();
                    workSheet.get_Range("D" + current_rownum_1, "D" + (current_rownum_1 + 7)).BorderAround2();
                    Microsoft.Office.Interop.Excel.Range oRange = (Microsoft.Office.Interop.Excel.Range)workSheet.Cells[current_rownum_1, 4];
                    float Left = (float)((double)oRange.Left);
                    float Top = (float)((double)oRange.Top);
                    const float ImageSize = 34;
                    workSheet.Shapes.AddPicture("https://chart.googleapis.com/chart?chs=150x150&cht=qr&chl=https://ams.fdcc.vn/Thiet_Bi/Edit/" + DataRow[i][0], MsoTriState.msoFalse, MsoTriState.msoCTrue, Left + 3, Top + 1, ImageSize, ImageSize);

                    current_rownum_1 = current_rownum_1 + 3;

                    foreach (Microsoft.Office.Interop.Excel.Range cell in workSheet.get_Range("B" + current_rownum_1, "C" + (current_rownum_1 + 4)))
                    {
                        cell.BorderAround2();
                    }

                    oSheet.Cells[current_rownum_1, 2]   = "Tên Thiết bị:";
                    oSheet.Cells[current_rownum_1, 2].Font.Bold = true;
                    oSheet.Cells[current_rownum_1, 3] = DataRow[i][1].ToString();
                    current_rownum_1++;

                    oSheet.Cells[current_rownum_1, 2] = "Ngày cấp:";
                    oSheet.Cells[current_rownum_1, 2].Font.Bold = true;
                    oSheet.Cells[current_rownum_1, 3] = "27-08-2018";
                    current_rownum_1++;

                    oSheet.Cells[current_rownum_1, 2] = "Phòng/Ban:";
                    oSheet.Cells[current_rownum_1, 2].Font.Bold = true;
                    oSheet.Cells[current_rownum_1, 3] = DataRow[i][2].ToString();
                    current_rownum_1++;

                    oSheet.Cells[current_rownum_1, 2] = "Group:";
                    oSheet.Cells[current_rownum_1, 2].Font.Bold = true;
                    oSheet.Cells[current_rownum_1, 3] = DataRow[i][13].ToString();
                    current_rownum_1++;

                    oSheet.Cells[current_rownum_1, 2] = "Mã Thiết bị:";
                    oSheet.Cells[current_rownum_1, 2].Font.Bold = true;
                    oSheet.Cells[current_rownum_1, 3] = DataRow[i][6].ToString();
                    current_rownum_1++;

                    if (((i != 0) && (((i+1) % 64) == 0)))
                    {
                        oSheet.Cells[current_rownum_1, 2].RowHeight = 10;
                        current_rownum_1 = current_rownum_1 + 2;
                        current_rownum_1 = current_rownum_1 + 3;
                        current_rownum_1++;
                        oSheet.Cells[current_rownum_1, 2].RowHeight = 2;
                    }
                    else
                    {
                        oSheet.Cells[current_rownum_1, 2].RowHeight = 10;
                    }
                    
                    //------------------------------QR_CARD_1------------------------------//
                }
                else if (i % 4 == 1)
                {
                    //------------------------------QR_CARD_2------------------------------//
                    current_rownum_2++;

                    workSheet.get_Range("F" + current_rownum_2, "F" + (current_rownum_2 + 2)).Merge();
                    workSheet.get_Range("F" + current_rownum_2, "F" + (current_rownum_2 + 2)).BorderAround2();
                    Microsoft.Office.Interop.Excel.Range oRange_logo = (Microsoft.Office.Interop.Excel.Range)workSheet.Cells[current_rownum_2, 6];
                    float Left_logo = (float)((double)oRange_logo.Left);
                    float Top_logo = (float)((double)oRange_logo.Top);
                    const float ImageSize_logo_W = 16;
                    const float ImageSize_logo_H = 9;
                    workSheet.Shapes.AddPicture(filepathImageLogo, MsoTriState.msoFalse, MsoTriState.msoCTrue, Left_logo + 4, Top_logo + 3, ImageSize_logo_W, ImageSize_logo_H);

                    workSheet.get_Range("G" + current_rownum_2, "G" + (current_rownum_2 + 2)).Merge();
                    workSheet.get_Range("G" + current_rownum_2, "G" + (current_rownum_2 + 2)).BorderAround2();
                    oSheet.Cells[current_rownum_2, 7] = "TEM THIẾT BỊ VĂN PHÒNG";
                    oSheet.Cells[current_rownum_2, 7].Font.Bold = true;

                    workSheet.get_Range("H" + current_rownum_2, "H" + (current_rownum_2 + 7)).Merge();
                    workSheet.get_Range("H" + current_rownum_2, "H" + (current_rownum_2 + 7)).BorderAround2();
                    Microsoft.Office.Interop.Excel.Range oRange = (Microsoft.Office.Interop.Excel.Range)workSheet.Cells[current_rownum_2, 8];
                    float Left = (float)((double)oRange.Left);
                    float Top = (float)((double)oRange.Top);
                    const float ImageSize = 34;
                    workSheet.Shapes.AddPicture("https://chart.googleapis.com/chart?chs=150x150&cht=qr&chl=https://ams.fdcc.vn/Thiet_Bi/Edit/" + DataRow[i][0], MsoTriState.msoFalse, MsoTriState.msoCTrue, Left + 3, Top + 1, ImageSize, ImageSize);

                    current_rownum_2 = current_rownum_2 + 3;

                    foreach (Microsoft.Office.Interop.Excel.Range cell in workSheet.get_Range("F" + current_rownum_2, "G" + (current_rownum_2 + 4)))
                    {
                        cell.BorderAround2();
                    }

                    oSheet.Cells[current_rownum_2, 6] = "Tên Thiết bị:";
                    oSheet.Cells[current_rownum_2, 6].Font.Bold = true;
                    oSheet.Cells[current_rownum_2, 7] = DataRow[i][1].ToString();
                    current_rownum_2++;

                    oSheet.Cells[current_rownum_2, 6] = "Ngày cấp:";
                    oSheet.Cells[current_rownum_2, 6].Font.Bold = true;
                    oSheet.Cells[current_rownum_2, 7] = "27-08-2018";
                    current_rownum_2++;

                    oSheet.Cells[current_rownum_2, 6] = "Phòng/Ban:";
                    oSheet.Cells[current_rownum_2, 6].Font.Bold = true;
                    oSheet.Cells[current_rownum_2, 7] = DataRow[i][2].ToString();
                    current_rownum_2++;

                    oSheet.Cells[current_rownum_2, 6] = "Group:";
                    oSheet.Cells[current_rownum_2, 6].Font.Bold = true;
                    oSheet.Cells[current_rownum_2, 7] = DataRow[i][13].ToString();
                    current_rownum_2++;

                    oSheet.Cells[current_rownum_2, 6] = "Mã Thiết bị:";
                    oSheet.Cells[current_rownum_2, 6].Font.Bold = true;
                    oSheet.Cells[current_rownum_2, 7] = DataRow[i][6];
                    current_rownum_2++;

                    if (((i != 0) && (((i + 1) % 64) == 0)))
                    {
                        oSheet.Cells[current_rownum_2, 6].RowHeight = 10;
                        current_rownum_2 = current_rownum_2 + 2;
                        current_rownum_2 = current_rownum_2 + 3;
                        current_rownum_2++;
                        oSheet.Cells[current_rownum_2, 6].RowHeight = 2;
                    }
                    else
                    {
                        oSheet.Cells[current_rownum_2, 6].RowHeight = 10;
                    }

                    //------------------------------QR_CARD_2------------------------------//
                }
                else if (i % 4 == 2)
                {
                    //------------------------------QR_CARD_3------------------------------//
                    current_rownum_3++;

                    workSheet.get_Range("J" + current_rownum_3, "J" + (current_rownum_3 + 2)).Merge();
                    workSheet.get_Range("J" + current_rownum_3, "J" + (current_rownum_3 + 2)).BorderAround2();
                    Microsoft.Office.Interop.Excel.Range oRange_logo = (Microsoft.Office.Interop.Excel.Range)workSheet.Cells[current_rownum_3, 10];
                    float Left_logo = (float)((double)oRange_logo.Left);
                    float Top_logo = (float)((double)oRange_logo.Top);
                    const float ImageSize_logo_W = 16;
                    const float ImageSize_logo_H = 9;
                    workSheet.Shapes.AddPicture(filepathImageLogo, MsoTriState.msoFalse, MsoTriState.msoCTrue, Left_logo + 4, Top_logo + 3, ImageSize_logo_W, ImageSize_logo_H);

                    workSheet.get_Range("K" + current_rownum_3, "K" + (current_rownum_3 + 2)).Merge();
                    workSheet.get_Range("K" + current_rownum_3, "K" + (current_rownum_3 + 2)).BorderAround2();
                    oSheet.Cells[current_rownum_3, 11] = "TEM THIẾT BỊ VĂN PHÒNG";
                    oSheet.Cells[current_rownum_3, 11].Font.Bold = true;

                    workSheet.get_Range("L" + current_rownum_3, "L" + (current_rownum_3 + 7)).Merge();
                    workSheet.get_Range("L" + current_rownum_3, "L" + (current_rownum_3 + 7)).BorderAround2();
                    Microsoft.Office.Interop.Excel.Range oRange = (Microsoft.Office.Interop.Excel.Range)workSheet.Cells[current_rownum_3, 12];
                    float Left = (float)((double)oRange.Left);
                    float Top = (float)((double)oRange.Top);
                    const float ImageSize = 34;
                    workSheet.Shapes.AddPicture("https://chart.googleapis.com/chart?chs=150x150&cht=qr&chl=https://ams.fdcc.vn/Thiet_Bi/Edit/" + DataRow[i][0], MsoTriState.msoFalse, MsoTriState.msoCTrue, Left + 3, Top + 1, ImageSize, ImageSize);

                    current_rownum_3 = current_rownum_3 + 3;

                    foreach (Microsoft.Office.Interop.Excel.Range cell in workSheet.get_Range("J" + current_rownum_3, "K" + (current_rownum_3 + 4)))
                    {
                        cell.BorderAround2();
                    }

                    oSheet.Cells[current_rownum_3, 10] = "Tên Thiết bị:";
                    oSheet.Cells[current_rownum_3, 10].Font.Bold = true;
                    oSheet.Cells[current_rownum_3, 11] = DataRow[i][1].ToString();
                    current_rownum_3++;

                    oSheet.Cells[current_rownum_3, 10] = "Ngày cấp:";
                    oSheet.Cells[current_rownum_3, 10].Font.Bold = true;
                    oSheet.Cells[current_rownum_3, 11] = "27-08-2018";
                    current_rownum_3++;

                    oSheet.Cells[current_rownum_3, 10] = "Phòng/Ban:";
                    oSheet.Cells[current_rownum_3, 10].Font.Bold = true;
                    oSheet.Cells[current_rownum_3, 11] = DataRow[i][2].ToString();
                    current_rownum_3++;

                    oSheet.Cells[current_rownum_3, 10] = "Group:";
                    oSheet.Cells[current_rownum_3, 10].Font.Bold = true;
                    oSheet.Cells[current_rownum_3, 11] = DataRow[i][13].ToString();
                    current_rownum_3++;

                    oSheet.Cells[current_rownum_3, 10] = "Mã Thiết bị:";
                    oSheet.Cells[current_rownum_3, 10].Font.Bold = true;
                    oSheet.Cells[current_rownum_3, 11] =  DataRow[i][6].ToString();
                    current_rownum_3++;

                    if (((i != 0) && (((i + 1) % 64) == 0)))
                    {
                        oSheet.Cells[current_rownum_3, 10].RowHeight = 10;
                        current_rownum_3 = current_rownum_3 + 2;
                        current_rownum_3 = current_rownum_3 + 3;
                        current_rownum_3++;
                        oSheet.Cells[current_rownum_3, 10].RowHeight = 2;
                    }
                    else
                    {
                        oSheet.Cells[current_rownum_3, 10].RowHeight = 10;
                    }
                    //------------------------------QR_CARD_3------------------------------//
                }
                else if (i % 4 == 3)
                {
                    //------------------------------QR_CARD_4------------------------------//
                    current_rownum_4++;

                    workSheet.get_Range("N" + current_rownum_4, "N" + (current_rownum_4 + 2)).Merge();
                    workSheet.get_Range("N" + current_rownum_4, "N" + (current_rownum_4 + 2)).BorderAround2();
                    Microsoft.Office.Interop.Excel.Range oRange_logo = (Microsoft.Office.Interop.Excel.Range)workSheet.Cells[current_rownum_4, 14];
                    float Left_logo = (float)((double)oRange_logo.Left);
                    float Top_logo = (float)((double)oRange_logo.Top);
                    const float ImageSize_logo_W = 16;
                    const float ImageSize_logo_H = 9;
                    workSheet.Shapes.AddPicture(filepathImageLogo, MsoTriState.msoFalse, MsoTriState.msoCTrue, Left_logo + 4, Top_logo + 3, ImageSize_logo_W, ImageSize_logo_H);

                    workSheet.get_Range("O" + current_rownum_4, "O" + (current_rownum_4 + 2)).Merge();
                    workSheet.get_Range("O" + current_rownum_4, "O" + (current_rownum_4 + 2)).BorderAround2();
                    oSheet.Cells[current_rownum_4, 15] = "TEM THIẾT BỊ VĂN PHÒNG";
                    oSheet.Cells[current_rownum_4, 15].Font.Bold = true;

                    workSheet.get_Range("P" + current_rownum_4, "P" + (current_rownum_4 + 7)).Merge();
                    workSheet.get_Range("P" + current_rownum_4, "P" + (current_rownum_4 + 7)).BorderAround2();
                    Microsoft.Office.Interop.Excel.Range oRange = (Microsoft.Office.Interop.Excel.Range)workSheet.Cells[current_rownum_4, 16];
                    float Left = (float)((double)oRange.Left);
                    float Top = (float)((double)oRange.Top);
                    const float ImageSize = 34;
                    workSheet.Shapes.AddPicture("https://chart.googleapis.com/chart?chs=150x150&cht=qr&chl=https://ams.fdcc.vn/Thiet_Bi/Edit/" + DataRow[i][0], MsoTriState.msoFalse, MsoTriState.msoCTrue, Left + 3, Top + 1, ImageSize, ImageSize);

                    current_rownum_4 = current_rownum_4 + 3;

                    foreach (Microsoft.Office.Interop.Excel.Range cell in workSheet.get_Range("N" + current_rownum_4, "O" + (current_rownum_4 + 4)))
                    {
                        cell.BorderAround2();
                    }

                    oSheet.Cells[current_rownum_4, 14] = "Tên Thiết bị:";
                    oSheet.Cells[current_rownum_4, 14].Font.Bold = true;
                    oSheet.Cells[current_rownum_4, 15] = DataRow[i][1].ToString();
                    current_rownum_4++;

                    oSheet.Cells[current_rownum_4, 14] = "Ngày cấp:";
                    oSheet.Cells[current_rownum_4, 14].Font.Bold = true;
                    oSheet.Cells[current_rownum_4, 15] = "27-08-2018";
                    current_rownum_4++;

                    oSheet.Cells[current_rownum_4, 14] = "Phòng/Ban:";
                    oSheet.Cells[current_rownum_4, 14].Font.Bold = true;
                    oSheet.Cells[current_rownum_4, 15] = DataRow[i][2].ToString();
                    current_rownum_4++;

                    oSheet.Cells[current_rownum_4, 14] = "Group:";
                    oSheet.Cells[current_rownum_4, 14].Font.Bold = true;
                    oSheet.Cells[current_rownum_4, 15] = DataRow[i][13].ToString();
                    current_rownum_4++;

                    oSheet.Cells[current_rownum_4, 14] = "Mã Thiết bị:";
                    oSheet.Cells[current_rownum_4, 14].Font.Bold = true;
                    oSheet.Cells[current_rownum_4, 15] = DataRow[i][6].ToString();
                    current_rownum_4++;

                    if (((i != 0) && (((i + 1) % 64) == 0)))
                    {
                        //oSheet.Cells[current_rownum_1, 2].RowHeight = 10;
                        current_rownum_1 = current_rownum_1 + 2;
                        current_rownum_1 = current_rownum_1 + 3;
                        current_rownum_1++;
                        oSheet.Cells[current_rownum_1, 2].RowHeight = 2;

                        //oSheet.Cells[current_rownum_2, 6].RowHeight = 10;
                        current_rownum_2 = current_rownum_2 + 2;
                        current_rownum_2 = current_rownum_2 + 3;
                        current_rownum_2++;
                        oSheet.Cells[current_rownum_2, 6].RowHeight = 2;

                        //oSheet.Cells[current_rownum_3, 10].RowHeight = 10;
                        current_rownum_3 = current_rownum_3 + 2;
                        current_rownum_3 = current_rownum_3 + 3;
                        current_rownum_3++;
                        oSheet.Cells[current_rownum_3, 10].RowHeight = 2;

                        oSheet.Cells[current_rownum_4, 14].RowHeight = 10;
                        current_rownum_4 = current_rownum_4 + 2;
                        current_rownum_4 = current_rownum_4 + 3;
                        current_rownum_4++;
                        oSheet.Cells[current_rownum_4, 14].RowHeight = 2;
                    }
                    else
                    {
                        oSheet.Cells[current_rownum_4, 14].RowHeight = 10;
                    }
                    //------------------------------QR_CARD_4------------------------------//
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

        public void Excel_Export_Large_Template(int Phong_Ban, int Group_Code)
        {
            List<int> Section_RowNum = new List<int>();

            int current_rownum_right = 1;
            int current_rownum_mid = 1;
            int current_rownum_left = 1;
            int Card_number;
            ProjectViewModel model = new ProjectViewModel();

            //using (OnlineShopDbContext db = new OnlineShopDbContext())
            //{
                //--------Add Dropdown for Type-------------------//
            //    model.Thiet_Bi = db.Thiet_Bis.OrderBy(m => m.ID).ToList();
            //    Card_number = db.Thiet_Bis.OrderBy(m => m.ID).Count();
            //}
            model.Thiet_Bi_Table = Load_LLTC_Excel_Report_By_Condition(Phong_Ban, Group_Code);
            Card_number = model.Thiet_Bi_Table.Rows.Count;

            DataRow[] DataRow = model.Thiet_Bi_Table.Select();

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
                    workSheet.Shapes.AddPicture("https://chart.googleapis.com/chart?chs=150x150&cht=qr&chl=https://ams.fdcc.vn/Thiet_Bi/Edit/" + DataRow[i][0], MsoTriState.msoFalse, MsoTriState.msoCTrue, Left + 4, Top + 10, ImageSize, ImageSize);

                    foreach (Microsoft.Office.Interop.Excel.Range cell in workSheet.get_Range("A" + current_rownum_right, "B" + (current_rownum_right + 4)))
                    {
                        cell.BorderAround2();
                    }

                    oSheet.Cells[current_rownum_right, 1] = "Tên Thiết bị:";
                    oSheet.Cells[current_rownum_right, 1].Font.Bold = true;
                    oSheet.Cells[current_rownum_right, 2] = DataRow[i][1].ToString();
                    current_rownum_right++;

                    oSheet.Cells[current_rownum_right, 1] = "Ngày cấp:";
                    oSheet.Cells[current_rownum_right, 1].Font.Bold = true;
                    oSheet.Cells[current_rownum_right, 2] = "27-08-2018";
                    current_rownum_right++;

                    oSheet.Cells[current_rownum_right, 1] = "Phòng/Ban:";
                    oSheet.Cells[current_rownum_right, 1].Font.Bold = true;
                    oSheet.Cells[current_rownum_right, 2] = DataRow[i][2].ToString();
                    current_rownum_right++;

                    oSheet.Cells[current_rownum_right, 1] = "Group:";
                    oSheet.Cells[current_rownum_right, 1].Font.Bold = true;
                    oSheet.Cells[current_rownum_right, 2] = DataRow[i][13].ToString();
                    current_rownum_right++;

                    oSheet.Cells[current_rownum_right, 1] = "Mã Thiết bị:";
                    oSheet.Cells[current_rownum_right, 1].Font.Bold = true;
                    oSheet.Cells[current_rownum_right, 2] = DataRow[i][6].ToString();
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
                    workSheet.Shapes.AddPicture("https://chart.googleapis.com/chart?chs=150x150&cht=qr&chl=https://ams.fdcc.vn/Thiet_Bi/Edit/" + DataRow[i][0], MsoTriState.msoFalse, MsoTriState.msoCTrue, Left + 4, Top + 10, ImageSize, ImageSize);

                    foreach (Microsoft.Office.Interop.Excel.Range cell in workSheet.get_Range("E" + current_rownum_mid, "F" + (current_rownum_mid + 4)))
                    {
                        cell.BorderAround2();
                    }

                    oSheet.Cells[current_rownum_mid, 5] = "Tên Thiết bị:";
                    oSheet.Cells[current_rownum_mid, 5].Font.Bold = true;
                    oSheet.Cells[current_rownum_mid, 6] = DataRow[i][1].ToString();
                    current_rownum_mid++;

                    oSheet.Cells[current_rownum_mid, 5] = "Ngày cấp:";
                    oSheet.Cells[current_rownum_mid, 5].Font.Bold = true;
                    oSheet.Cells[current_rownum_mid, 6] = "27-08-2018";
                    current_rownum_mid++;

                    oSheet.Cells[current_rownum_mid, 5] = "Phòng/Ban:";
                    oSheet.Cells[current_rownum_mid, 5].Font.Bold = true;
                    oSheet.Cells[current_rownum_mid, 6] = DataRow[i][2].ToString();
                    current_rownum_mid++;

                    oSheet.Cells[current_rownum_mid, 5] = "Group:";
                    oSheet.Cells[current_rownum_mid, 5].Font.Bold = true;
                    oSheet.Cells[current_rownum_mid, 6] = DataRow[i][13].ToString();
                    current_rownum_mid++;

                    oSheet.Cells[current_rownum_mid, 5] = "Mã Thiết bị:";
                    oSheet.Cells[current_rownum_mid, 5].Font.Bold = true;
                    oSheet.Cells[current_rownum_mid, 6] = DataRow[i][6].ToString();
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
                    workSheet.Shapes.AddPicture("https://chart.googleapis.com/chart?chs=150x150&cht=qr&chl=https://ams.fdcc.vn/Thiet_Bi/Edit/" + DataRow[i][0], MsoTriState.msoFalse, MsoTriState.msoCTrue, Left + 4, Top + 10, ImageSize, ImageSize);

                    foreach (Microsoft.Office.Interop.Excel.Range cell in workSheet.get_Range("I" + current_rownum_left, "J" + (current_rownum_left + 4)))
                    {
                        cell.BorderAround2();
                    }

                    oSheet.Cells[current_rownum_left, 9] = "Tên Thiết bị:";
                    oSheet.Cells[current_rownum_left, 9].Font.Bold = true;
                    oSheet.Cells[current_rownum_left, 10] = DataRow[i][1].ToString();
                    current_rownum_left++;

                    oSheet.Cells[current_rownum_left, 9] = "Ngày cấp:";
                    oSheet.Cells[current_rownum_left, 9].Font.Bold = true;
                    oSheet.Cells[current_rownum_left, 10] = "27-08-2018";
                    current_rownum_left++;

                    oSheet.Cells[current_rownum_left, 9] = "Phòng/Ban:";
                    oSheet.Cells[current_rownum_left, 9].Font.Bold = true;
                    oSheet.Cells[current_rownum_left, 10] = DataRow[i][2].ToString();
                    current_rownum_left++;

                    oSheet.Cells[current_rownum_left, 9] = "Group:";
                    oSheet.Cells[current_rownum_left, 9].Font.Bold = true;
                    oSheet.Cells[current_rownum_left, 10] = DataRow[i][13].ToString();
                    current_rownum_left++;

                    oSheet.Cells[current_rownum_left, 9] = "Mã Thiết bị:";
                    oSheet.Cells[current_rownum_left, 9].Font.Bold = true;
                    oSheet.Cells[current_rownum_left, 10] = DataRow[i][6].ToString();
                    current_rownum_left++;

                    if (((i != 0) && (((i + 1) % 21) == 0)))
                    {
                        oSheet.Cells[current_rownum_right, 1].RowHeight = 13;
                        current_rownum_right++;
                        oSheet.Cells[current_rownum_mid, 5].RowHeight = 13;
                        current_rownum_mid++;
                        oSheet.Cells[current_rownum_left, 9].RowHeight = 13;
                        current_rownum_left++;
                    }
                    else
                    {
                        oSheet.Cells[current_rownum_left, 9].RowHeight = 24;
                    }

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
        public ActionResult Interop_Index(ProjectViewModel GetIndex)
        {

            ProjectViewModel model = new ProjectViewModel();

            model = GetIndex;

            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                model.Thiet_Bi_Table = Load_LLTC_Excel_Report_By_Condition(GetIndex.Select_Phong_Ban, GetIndex.Select_Group);
                //--------Add Dropdown for Code_Group-------------------//
                model.Code_Group = db.Code_Group.OrderBy(m => m.ID).ToList();
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

                //--------Add Dropdown for Phong_Ban-------------------//
                model.CS_tbPhong_Ban = db.CS_tbPhong_Ban.OrderBy(m => m.ID).ToList();
                model.Phong_Ban_All = new List<SelectListItem>();
                var items_2 = new List<SelectListItem>();
                foreach (var Phong_ban_Main in model.CS_tbPhong_Ban)
                {
                    items_2.Add(new SelectListItem()
                    {
                        Value = Phong_ban_Main.ID.ToString(),
                        Text = Phong_ban_Main.Type
                    });
                }
                model.Phong_Ban_All = items_2;
                //--------Add Dropdown for Phong_Ban-------------------//
            }

            return View("Index_2", model);

        }

        [HttpPost]
        public ActionResult Interop_Index_Main(ProjectViewModel GetIndex)
        {

            ProjectViewModel model = new ProjectViewModel();

            model = GetIndex;

            using (OnlineShopDbContext db = new OnlineShopDbContext())
            {
                model.Thiet_Bi_Table = Load_LLTC_Excel_Report_By_Condition(GetIndex.Select_Phong_Ban, GetIndex.Select_Group);
                //--------Add Dropdown for Code_Group-------------------//
                model.Code_Group = db.Code_Group.OrderBy(m => m.ID).ToList();
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

                //--------Add Dropdown for Phong_Ban-------------------//
                model.CS_tbPhong_Ban = db.CS_tbPhong_Ban.OrderBy(m => m.ID).ToList();
                model.Phong_Ban_All = new List<SelectListItem>();
                var items_2 = new List<SelectListItem>();
                foreach (var Phong_ban_Main in model.CS_tbPhong_Ban)
                {
                    items_2.Add(new SelectListItem()
                    {
                        Value = Phong_ban_Main.ID.ToString(),
                        Text = Phong_ban_Main.Type
                    });
                }
                model.Phong_Ban_All = items_2;
                //--------Add Dropdown for Phong_Ban-------------------//
            }

            return View("Index", model);

        }

        [HttpPost]
        public ActionResult Interop(ProjectViewModel model)
        {
            if (model.Select_Size == 0)
            {
                Excel_Export_Large_Template(model.Select_Phong_Ban, model.Select_Group);
            }
            else
            {
                Excel_Export_Small_Template(model.Select_Phong_Ban, model.Select_Group);
            }
            
            return RedirectToAction("Interop_Index", "Thiet_Bi", model);

        }

        System.Data.DataTable Load_LLTC_Excel_Report_By_Condition(int Phong_Ban, int Ma_Nhom)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            SqlConnection conn = null;
            conn = new SqlConnection(string.Format("Data Source=SRBDC.FDC.LOCAL; Initial Catalog=EQUIP; User id=sa; Password=P@ssw0rd"));
            try
            {
                cmd = new SqlCommand("Thiet_Bi_List_By_Condition", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Phong_Ban", Phong_Ban);
                cmd.Parameters.AddWithValue("@Ma_Nhom", Ma_Nhom);
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

        System.Data.DataTable Load_ThietBi()
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            SqlConnection conn = null;
            conn = new SqlConnection(string.Format("Data Source=SRBDC.FDC.LOCAL; Initial Catalog=EQUIP; User id=sa; Password=P@ssw0rd"));
            try
            {
                cmd = new SqlCommand("Thiet_Bi_List", conn);
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
