using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsApp1.Tables
{
    public class FASON_İMALAT_EXCEL
    {
        [Description("Item")]
        public string Item { get; set; }

        [Description("Part Number")]
        public string Part_Number { get; set; }

        [Description("Item QTY")]
        public string Item_QTY { get; set; }

        [Description("Module No")]
        public string Module_No { get; set; }
    }
}
