namespace Models.Framework
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class Code_Equip
    {
        public int ID { get; set; }

        public int ID_Code { get; set; }

        public string Equip { get; set; }
    }
}
