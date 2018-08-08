namespace Models.Framework
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class CS_tbWorkCount_Sub
    {
        public int ID { get; set; }

        public int? CS_tbWorkCount_ID { get; set; }

        public int? CS_tbLLTCTypeSub_ID { get; set; }

        public int? CS_LLTC_ID { get; set; }

        [Required(ErrorMessage = "Không được để trống nội dung này")]
        public int? CS_tbNumberDailyCount { get; set; }
    }
}
