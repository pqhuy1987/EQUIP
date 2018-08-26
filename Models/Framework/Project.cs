namespace Models.Framework
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Thiet_Bi")]
    public partial class Thiet_Bi
    {
        public int ID { get; set; }

        //[Required(ErrorMessage = "Không được để trống nội dung này")]
        //[StringLength(50, ErrorMessage = "Nội dung nhập vào không quá 50 ký tự")]
        public string Ten_Thiet_Bi { get; set; }

        //[Required(ErrorMessage = "Không được để trống nội dung này")]
        //[StringLength(50, ErrorMessage = "Nội dung nhập vào không quá 50 ký tự")]
        public int Phong_Ban { get; set; }

        //[Required(ErrorMessage = "Không được để trống nội dung này")]
        //[StringLength(50, ErrorMessage = "Nội dung nhập vào không quá 50 ký tự")]
        public int Ma_Nhom { get; set; }

        //[Required(ErrorMessage = "Không được để trống nội dung này")]
        //[StringLength(50, ErrorMessage = "Nội dung nhập vào không quá 50 ký tự")]
        public int Ma_Chi_Tiet { get; set; }

        //[Required(ErrorMessage = "Không được để trống nội dung này")]
        //[StringLength(50, ErrorMessage = "Nội dung nhập vào không quá 50 ký tự")]
        public int Vi_Tri    { get; set; }

        //[Required(ErrorMessage = "Không được để trống nội dung này")]
        //[StringLength(50, ErrorMessage = "Nội dung nhập vào không quá 50 ký tự")]
        public string Hinh_Anh { get; set; }

        public string Ma_Thiet_Bi { get; set; }

        //[Required(ErrorMessage = "Không được để trống nội dung này")]
        //[StringLength(50, ErrorMessage = "Nội dung nhập vào không quá 50 ký tự")]
        public string Ghi_Chu_1 { get; set; }

        //[Required(ErrorMessage = "Không được để trống nội dung này")]
        [Column(TypeName = "date")]
        [DataType(DataType.Date)]
        public DateTime? Start_Date { get; set; }

        //[Required(ErrorMessage = "Không được để trống nội dung này")]
        [DataType(DataType.Date)]
        [Column(TypeName = "date")]
        public DateTime? End_Date { get; set; }

        //[Required(ErrorMessage = "Không được để trống nội dung này")]
        //[StringLength(50, ErrorMessage = "Nội dung nhập vào không quá 50 ký tự")]
        public string Ghi_Chu_2 { get; set; }

        //[Required(ErrorMessage = "Không được để trống nội dung này")]
        //[StringLength(50, ErrorMessage = "Nội dung nhập vào không quá 50 ký tự")]
        public string Don_Gia { get; set; }

        public int Number_Person;
        public int Number_Project;
    }
}
