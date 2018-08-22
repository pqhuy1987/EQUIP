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
    public class Code_EquipViewModel
    {
        public List<Code_Equip> Code_Equip { get; set; }
        public Code_Equip Code_EquipSelect { get; set; }

        public List<Code_Group> Code_Group { get; set; }
        public List<SelectListItem> Code_Group_All { get; set; }
    }
}
