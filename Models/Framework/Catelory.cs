namespace Models.Framework
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Catelory")]
    public partial class Catelory
    {
        public int ID { get; set; }

        [StringLength(100)]
        public string Prj_Name { get; set; }

        [StringLength(100)]
        public string Unit_Name { get; set; }

        [StringLength(100)]
        public string Owner_Name { get; set; }

        [StringLength(50)]
        public string Phone_Number { get; set; }

        public int? Person_Number { get; set; }

        public DateTime? Create_Date { get; set; }

        [StringLength(50)]
        public string Status { get; set; }

        [StringLength(50)]
        public string Email { get; set; }

        [StringLength(100)]
        public string Area { get; set; }

        [StringLength(100)]
        public string Rate { get; set; }

        [StringLength(100)]
        public string Job { get; set; }
    }
}
