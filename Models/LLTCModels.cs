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
    public class LLTCViewModel
    {
        public List<LLTC>                   LLTC                    { get; set; }
        public LLTC                         SelectedLLTC            { get; set; }

        public List<CS_tbLLTCTypeSub>       CS_tbLLTCTypeSub        { get; set; }
        public CS_tbLLTCTypeSub             CS_tbLLTCTypeSub_Select { get; set; }

        public List<Project>                Project                 { get; set; }
        public List<CS_tbWorkType>          CS_tbWorkType           { get; set; }
        public List<CS_tbViTri>      CS_tbViTri       { get; set; }

        public List<SelectListItem>         Project_Name_All        { get; set; }
        public List<SelectListItem>         WorkTypeDetails_All     { get; set; }
        public List<SelectListItem>         WorkTypeMain_All        { get; set; }

        public string                       DisplayMode             { get; set; }
    }

    class LLTCModels
    {
        private OnlineShopDbContext context = null;

        public LLTCModels()
        {
            context = new OnlineShopDbContext();
        }
    }
}