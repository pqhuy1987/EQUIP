using System;
using System.Data.SqlClient; //add for SqlParameter pqhuy1987
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Models.Framework;
using System.Web.Mvc;

namespace Models
{
    public class WorkCountViewModel
    {
        public List<DateTime>   SelectDate              { get; set; }
        public List<DateTime>   SelectDate_Temp         { get; set; }
        public List<Catelory>   Catelory                { get; set; }
        public List<Project>    Project                 { get; set; }
        public List<WorkCount> WorkCount                { get; set; }
        public List<List<WorkCount>>  WorkCount_List    { get; set; }
        public List<List<WorkCount>>  WorkCount_List_2  { get; set; }
        public List<WorkCount>  WorkCount_Temp          { get; set; }
        public List<WorkCount>  WorkCount_Temp_2        { get; set; }

        public List<int>  Count_Number                   { get; set; }

        public Project SelectedProject                  { get; set; }
        public WorkCount SelectedWorkCount              { get; set; }

        public List<Catelory> Catelory_Project          { get; set; }
        public List<SelectListItem> ProjectAll          { get; set; }

        public string DisplayMode   { get; set; }
        public DateTime StartDate   { get; set; }
        public DateTime EndDate     { get; set; }
        public int Number_Team_1    { get; set; }
        public int Number_Team_2    { get; set; }

        public List<string> List_Job    { get; set;}
        public int Number_Job           { get; set;}

        public List<int> Total_number   { get; set;}

        public List<List<int>> Total_Job_number     { get; set; }
        public List<int>       Total_Job_number_2   { get; set; }

        //Để tạm
        //public int Number_Person;
        //public int Number_Project;
    }

    class WorkCountModels
    {

    }
}
