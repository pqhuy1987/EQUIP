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
    public class CS_tbWorkCountViewModels
    {
        public List<CS_tbWorkCount>     CS_tbWorkCount                { get; set; }
        public List<CS_tbWorkCount_Sub> CS_tbWorkCount_Sub            { get; set; }
        public CS_tbWorkCount           CS_tbWorkCount_Select         { get; set; }
        public List<CS_tbWorkCount_Sub> CS_tbWorkCountMultiSelect_Sub { get; set; }

        public List<Project>            Project                       { get; set; }
        public List<CS_tbLLTCTypeSub>   CS_tbLLTCTypeSub              { get; set; }
        public List<LLTC>               LLTC_temp                     { get; set; }
        public List<CS_tbWorkType>      CS_tbWorkType_temp            { get; set; }
        public CS_tbLLTCTypeSub         CS_tbLLTCTypeSub_Select       { get; set; }

        public List<SelectListItem>     Project_Name_All              { get; set; }

        public string                   ValidStatus                   { get; set; }
    }
}
