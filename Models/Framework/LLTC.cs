namespace Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("LLTC")]
    public partial class LLTC
    {
        public int ID { get; set; }

        [Required(ErrorMessage = "Không được để trống nội dung này")]
        [StringLength(50, ErrorMessage = "Nội dung nhập vào không quá 50 ký tự")]
        public string Main_Name_LLTC { get; set; }

        [Required(ErrorMessage = "Không được để trống nội dung này")]
        [StringLength(50, ErrorMessage = "Nội dung nhập vào không quá 50 ký tự")]
        public string Main_Name_Ower { get; set; }

        [Required(ErrorMessage = "Không được để trống nội dung này")]
        public int Main_Name_Job { get; set; }

        [Required(ErrorMessage = "Không được để trống nội dung này")]
        [StringLength(50, ErrorMessage = "Nội dung nhập vào không quá 50 ký tự")]
        public string Main_Number { get; set; }

        [Required(ErrorMessage = "Không được để trống nội dung này")]
        [StringLength(50, ErrorMessage = "Nội dung nhập vào không quá 50 ký tự")]
        public string Main_Area { get; set; }

        [Required(ErrorMessage = "Không được để trống nội dung này")]
        [StringLength(50, ErrorMessage = "Nội dung nhập vào không quá 50 ký tự")]
        public string Main_Total_Number { get; set; }

        [Required(ErrorMessage = "Không được để trống nội dung này")]
        [StringLength(50, ErrorMessage = "Nội dung nhập vào không quá 50 ký tự")]
        public string Main_Rate { get; set; }

        [Required(ErrorMessage = "Không được để trống nội dung này")]
        [StringLength(100, ErrorMessage = "Nội dung nhập vào không quá 50 ký tự")]
        public string Main_Note { get; set; }

        [Required(ErrorMessage = "Không được để trống nội dung này")]
        [StringLength(50, ErrorMessage = "Nội dung nhập vào không quá 50 ký tự")]
        public string Main_Status { get; set; }
    }
}
