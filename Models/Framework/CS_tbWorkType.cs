namespace Models.Framework
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class CS_tbWorkType
    {
        public int ID { get; set; }

        [Required(ErrorMessage = "Không được để trống nội dung này")]
        public int CoreWorkType { get; set; }

        [Required(ErrorMessage = "Không được để trống nội dung này")]
        [StringLength(50, ErrorMessage = "Nội dung nhập vào không quá 50 ký tự")]
        public string SubWorkType { get; set; }
    }
}
