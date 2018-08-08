namespace Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class CS_tbWorkCount
    {
        public int ID { get; set; }

        [StringLength(50)]
        public string tb_WorkCountName_Report { get; set; }

        [Required(ErrorMessage = "Không được để trống nội dung này")]
        public int tb_WorkCountProject_ID { get; set; }

        [Required(ErrorMessage = "Không được để trống nội dung này")]
        public DateTime? tb_WorkCountForDate { get; set; }

        public DateTime? tb_WorkCountDateTime_Report { get; set; }

        public string tb_WorkCountName_Edit { get; set; }

        public DateTime? tb_WorkCountDateTime_Edit { get; set; }

        public int? tb_mTotalCount { get; set; }
    }
}
