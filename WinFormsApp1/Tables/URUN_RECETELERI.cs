using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsApp1.Tables
{
    public class URUN_RECETELERI
    {
        public Guid rec_Guid {  get; set; }
        public string rec_anakod { get; set;}
        public string rec_tuketim_kod {  get; set;}
        public double rec_tuketim_miktar { get; set;}
    }
}
