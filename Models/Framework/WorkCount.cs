namespace Models.Framework
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("WorkCount")]
    public partial class WorkCount
    {
        public int ID { get; set; }

        [Column(TypeName = "date")]
        public DateTime? CreateDate { get; set; }

        public string Ten_Thiet_Bi { get; set; }

        public string Unit_Name { get; set; }

        public int? Unit_Number { get; set; }

        [StringLength(50)]
        public string Unit_Job { get; set; }
    }
}
