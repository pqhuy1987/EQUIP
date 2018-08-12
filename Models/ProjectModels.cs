using System;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Models.Framework;
using System.Data.SqlClient;
using System.Web.Mvc;

namespace Models
{
    public class ProjectViewModel
    {
        public List<Project>                    Project                     { get; set; }
        public Project                          SelectedProject             { get; set; }
        public List<CS_tbLLTCTypeSub>           CS_tbLLTCTypeSub            { get; set; }
        public CS_tbLLTCTypeSub                 CS_tbLLTCTypeSub_Select     { get; set; }
        public List<LLTC>                       LLTC                        { get; set; }
        public LLTC                             LLTC_Select                 { get; set; }
        public List<CS_tbWorkType>              CS_tbWorkType               { get; set; }
        public List<CS_tbViTri>          CS_tbViTri           { get; set; }

        public string DisplayMode                                           { get; set; }
        public int DisplayModeSub                                           { get; set; }

        public List<CS_tbConstructionSiteType>  CS_tbConstructionSiteType   { get; set; }

        public List<SelectListItem>             Project_Type_All            { get; set; }
        public List<SelectListItem>             Project_All                 { get; set; }
        public List<SelectListItem>             LLTC_Name_All               { get; set; }
        public List<SelectListItem>             WorkTypeDetails_All         { get; set; }
        public List<SelectListItem>             WorkTypeCore_All            { get; set; }
    }

    public class ProjectModels
    {
        private OnlineShopDbContext context = null;

        public ProjectModels()
        {
            context = new OnlineShopDbContext();
        }

        public List<Project> ListAll()
        {
            var list = context.Database.SqlQuery<Project>("Sp_Project_ListAll").ToList();
            return list;
        }

        public int Create(string ProjectName)
        {
            object[] parameters =
            {
                new SqlParameter ("@ProjectName",ProjectName),

            };
            int res = context.Database.ExecuteSqlCommand("Sp_Project_Insert @ProjectName", parameters);
            return res;
        }
    }
}
