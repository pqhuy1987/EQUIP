using System;
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
using System.Runtime.InteropServices;

namespace Models
{
    public class ProjectViewModel
    {
        public DataTable                        Thiet_Bi_Table              { get; set; }
        public List<Thiet_Bi>                   Thiet_Bi                    { get; set; }
        public Thiet_Bi                         SelectedProject             { get; set; }
        public List<CS_tbLLTCTypeSub>           CS_tbLLTCTypeSub            { get; set; }
        public CS_tbLLTCTypeSub                 CS_tbLLTCTypeSub_Select     { get; set; }
        public List<LLTC>                       LLTC                        { get; set; }
        public LLTC                             LLTC_Select                 { get; set; }
        public LoginModel                       LoginModel                  { get; set; }
        public List<CS_tbWorkType>              CS_tbWorkType               { get; set; }
        public List<CS_tbViTri>                 CS_tbViTri                  { get; set; }

        public string DisplayMode                                           { get; set; }
        public int DisplayModeSub                                           { get; set; }

        public List<CS_tbPhong_Ban>             CS_tbPhong_Ban              { get; set; }

        public List<SelectListItem>             Phong_Ban_All               { get; set; }
        public List<SelectListItem>             Vi_Tri_All                  { get; set; }
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

        public List<Thiet_Bi> ListAll()
        {
            var list = context.Database.SqlQuery<Thiet_Bi>("Sp_Project_ListAll").ToList();
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
