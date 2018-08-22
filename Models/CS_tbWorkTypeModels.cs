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
    public class CS_tbWorkTypeViewModel
    {
        public List<CS_tbWorkType>      CS_tbWorkType { get; set; }
        public CS_tbWorkType            CS_tbWorkTypeSelect { get; set; }

        public List<CS_tbViTri>         CS_tbViTri { get; set; }
        public List<SelectListItem>     WorkTypeMain_All { get; set; }
    }
}
