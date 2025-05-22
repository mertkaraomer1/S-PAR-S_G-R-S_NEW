using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsApp1.Tables
{
    public class SATINALMA_TALEPLERI1
    {
        public Guid stl_Guid { get; set; } // uniqueidentifier
        public DateTime stl_create_date { get; set; } // datetime
        public string stl_evrak_seri { get; set; } // dbo.evrakseri_str (nullable)
        public int stl_evrak_sira { get; set; } // int (nullable)
        public string stl_Sor_Merk { get; set; } // nvarchar(25) (nullable)
        public string stl_Stok_kodu { get; set; } // nvarchar(25) (nullable)
        public string stl_HareketGrupKodu1 { get; set; } // nvarchar(25) (nullable)
        public DateTime stl_teslim_tarihi { get; set; }
        public double stl_miktari { get; set; }
        public double stl_teslim_miktari { get; set; }
        public bool? stl_kapat_fl { get; set; } // tinyint (nullable)
    }
}
