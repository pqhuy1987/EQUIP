namespace Models.Framework
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Phan_Quyen")]
    public partial class Phan_Quyen
    {
        public int ID { get; set; }

        [StringLength(50)]
        public string Dia_Chi_Mail { get; set; }

    }
}
