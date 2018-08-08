namespace Models.Framework
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class CS_tbLLTCTypeSub
    {
        public int ID { get; set; }

        public int? CS_tbLLTC_ID { get; set; }

        [StringLength(50)]
        public string CS_tbLLTCNameSub { get; set; }

        [Required(ErrorMessage = "Không được để trống nội dung này")]
        public int? CS_tbLLTCNameSiteID { get; set; }

        [Required(ErrorMessage = "Không được để trống nội dung này")]
        [StringLength(50, ErrorMessage = "Nội dung nhập vào không quá 50 ký tự")]
        public string CS_tbLLTCNumberRegisterSub { get; set; }

        [StringLength(50)]
        public string CS_tbLLTCNameSiteManagerSub { get; set; }

        [StringLength(50)]
        public string CS_tbLLTCNameSiteManagerMobileSub { get; set; }

        [Required(ErrorMessage = "Không được để trống nội dung này")]
        public int? CS_tbLLTCNameJobDetailsSub { get; set; }

        [Required(ErrorMessage = "Không được để trống nội dung này")]
        [DataType(DataType.Date)]
        [Column(TypeName = "date")]
        public DateTime? CS_tbLLTCStartDateSub { get; set; }

        [Required(ErrorMessage = "Không được để trống nội dung này")]
        [DataType(DataType.Date)]
        [Column(TypeName = "date")]
        public DateTime? CS_tbLLTCEndDateSub { get; set; }

        [Required(ErrorMessage = "Không được để trống nội dung này")]
        [StringLength(50, ErrorMessage = "Nội dung nhập vào không quá 50 ký tự")]
        public string CS_tbLLTCStatusSub { get; set; }
    }
}
