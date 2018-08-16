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
    public class CateloryViewModel
    {
        public List<Catelory>   Catelory            { get; set; }
        public List<Thiet_Bi>    Thiet_Bi             { get; set; }
        public List<LLTC>       LLTC                { get; set; }
        public List<Catelory>   Catelory_Project    { get; set; }

        public List<SelectListItem> ProjectAll  { get; set; }
        public List<SelectListItem> MainNameAll { get; set; }

        public Catelory SelectedCatelory { get; set; }

        public string DisplayMode { get; set; }
    }

    public class CateloryModels
    {
        private OnlineShopDbContext context = null;

        public CateloryModels()
        {
            context = new OnlineShopDbContext();
        }
    }
}
