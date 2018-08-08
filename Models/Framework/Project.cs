namespace Models.Framework
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Project")]
    public partial class Project
    {
        public int ID { get; set; }

        [Required(ErrorMessage = "Không được để trống nội dung này")]
        [StringLength(50, ErrorMessage = "Nội dung nhập vào không quá 50 ký tự")]
        public string Project_Name { get; set; }

        [Required(ErrorMessage = "Không được để trống nội dung này")]
        [StringLength(50, ErrorMessage = "Nội dung nhập vào không quá 50 ký tự")]
        public string Site_Type { get; set; }

        [Required(ErrorMessage = "Không được để trống nội dung này")]
        [StringLength(50, ErrorMessage = "Nội dung nhập vào không quá 50 ký tự")]
        public string General_Director { get; set; }

        [Required(ErrorMessage = "Không được để trống nội dung này")]
        [StringLength(50, ErrorMessage = "Nội dung nhập vào không quá 50 ký tự")]
        public string Site_Manager { get; set; }

        public string Site_Address { get; set; }

        [Required(ErrorMessage = "Không được để trống nội dung này")]
        [StringLength(50, ErrorMessage = "Nội dung nhập vào không quá 50 ký tự")]
        public string Value_Cost { get; set; }

        [Required(ErrorMessage = "Không được để trống nội dung này")]
        [Column(TypeName = "date")]
        [DataType(DataType.Date)]
        public DateTime? Start_Date { get; set; }

        [Required(ErrorMessage = "Không được để trống nội dung này")]
        [DataType(DataType.Date)]
        [Column(TypeName = "date")]
        public DateTime? End_Date { get; set; }

        [Required(ErrorMessage = "Không được để trống nội dung này")]
        [StringLength(50, ErrorMessage = "Nội dung nhập vào không quá 50 ký tự")]
        public string Operation_Status { get; set; }

        [Required(ErrorMessage = "Không được để trống nội dung này")]
        [StringLength(50, ErrorMessage = "Nội dung nhập vào không quá 50 ký tự")]
        public string Site_Area { get; set; }

        public int Number_Person;
        public int Number_Project;
    }
}
