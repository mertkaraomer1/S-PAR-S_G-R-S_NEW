using System.Data;
using WinFormsApp1.Context;
using Microsoft.Office.Interop.Excel;

namespace SİPARİŞ_GİRİŞ
{
    public partial class VERİLEN_TEKLİF_MALİYET_HESABI : Form
    {
        private MyDbContext dbContext;
        public VERİLEN_TEKLİF_MALİYET_HESABI()
        {
            dbContext = new MyDbContext();
            InitializeComponent();
        }
        System.Data.DataTable table = new System.Data.DataTable();
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            // DataGridView'deki verileri bir DataTable'a kopyalayın
            System.Data.DataTable dt = new System.Data.DataTable();

            foreach (DataGridViewColumn column in advancedDataGridView1.Columns)
            {
                // Eğer ValueType null ise, varsayılan bir veri türü kullanabilirsiniz.
                Type columnType = column.ValueType ?? typeof(string);
                dt.Columns.Add(column.HeaderText, columnType);
            }

            // Satırları ekle
            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                DataRow dataRow = dt.NewRow();
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Value != null) // Hücre değeri null değilse
                    {
                        dataRow[cell.ColumnIndex] = cell.Value;
                    }
                    else if (cell is DataGridViewCheckBoxCell) // CheckBox hücresi ise
                    {
                        dataRow[cell.ColumnIndex] = (cell.Value != null && (bool)cell.Value) ? "True" : "False"; // CheckBox değeri null değilse ve true ise "True" olarak ayarla, değilse "False" olarak ayarla
                    }
                }
                dt.Rows.Add(dataRow);
            }

            // Excel uygulamasını başlatın
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;

            // Yeni bir Excel çalışma kitabı oluşturun
            Workbook workbook = excelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            _Worksheet worksheet = (_Worksheet)workbook.Worksheets[1];

            // DataTable'ı Excel çalışma sayfasına aktarın
            int rowIndex = 1;

            // Başlıkları yaz
            for (int j = 0; j < dt.Columns.Count; j++)
            {
                worksheet.Cells[1, j + 1] = dt.Columns[j].ColumnName;
                worksheet.Cells[1, j + 1].Font.Bold = true;
            }

            // Verileri yaz
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                rowIndex++;
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    // Metin olarak aktar
                    worksheet.Cells[rowIndex, j + 1].NumberFormat = "@";
                    worksheet.Cells[rowIndex, j + 1] = dt.Rows[i][j].ToString();
                }
            }
        }

        private void VERİLEN_TEKLİF_MALİYET_HESABI_Load(object sender, EventArgs e)
        {
            toolStripLabel1.ToolTipText =

               "1.Evrak Sıra No Yaz\n" +

               "2.İşçilerin Dakikalık Ücretini Yaz\n" +

               "3.Listele";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            table.Columns.Clear();
            table.Rows.Clear();
            table.Clear();
            table.Columns.Add("SATIR_NO");
            table.Columns.Add("STOK KODU");
            table.Columns.Add("SİPARİŞ TARİHİ");
            table.Columns.Add("TÜKETİM KODU");
            table.Columns.Add("TÜKETİM KODU ADI");
            table.Columns.Add("TÜKETİM KODU KG MİKTAR");
            table.Columns.Add("BİRİM FİYATI");
            table.Columns.Add("PLANLANAN İŞÇİ MALİYETİ");
            table.Columns.Add("TAMAMLANAN İŞÇİ MALİYETİ");
            table.Columns.Add("TALEP MİKTARI");
            table.Columns.Add("SATINALMA FİYATI");
            table.Columns.Add("SATINALMA FİYATI TÜKETİM KOD");
            table.Columns.Add("AÇIKLAMA");

            string anakod = textBox1.Text;
            int DKIsciMaliyet = Convert.ToInt32(textBox2.Text);


            var Talepler =dbContext.URUN_RECETELERI.Where(x => x.rec_anakod == anakod)
                    .Select(x => new
                    {
                        x.rec_anakod,
                        x.rec_tuketim_kod,
                        x.rec_tuketim_miktar,
                        gereklimiktar = x.rec_tuketim_miktar * 1
                    }).Join(dbContext.STOKLAR,
                tuket => tuket.rec_tuketim_kod,
                stok => stok.sto_kod,
                    (tuket, stok) => new
                    {
                        tuket.rec_anakod,
                        tuket.rec_tuketim_kod,
                        tuket.gereklimiktar,
                        tuket.rec_tuketim_miktar,
                        stok.sto_isim
                    })
                .ToList();

            if (Talepler.Count() == 0)
            {
                int sayac = 1;
                var MalzemeSatisFiyatı = dbContext.SIPARISLER
                    .Where(x => x.sip_stok_kod == anakod && x.sip_evrakno_seri == "")
                    .OrderByDescending(x => x.sip_tarih).Select(x => new
                    {
                        x.sip_stok_kod,
                        x.sip_tarih,
                        x.sip_miktar,
                        x.sip_aciklama,
                        TopFiyatSatis = x.sip_b_fiyat * 1 * x.sip_doviz_kuru,

                    })
                    .FirstOrDefault();


                table.Rows.Add(
                         sayac++,
                         anakod,
                         MalzemeSatisFiyatı?.sip_tarih.Date.ToString("yyyy/MM/dd"),
                         "",
                         "",
                         MalzemeSatisFiyatı?.sip_miktar ?? 0,
                         0,
                         0,
                         0,
                         "1",
                         Math.Round(MalzemeSatisFiyatı?.TopFiyatSatis ?? 0, 2),
                         0,
                         MalzemeSatisFiyatı?.sip_aciklama
                     );
                advancedDataGridView2.DataSource = table;
            }

            else if (Talepler.Count() != 0)
            {
                int sayac = 1;
                foreach (var item in Talepler)
                {
                    var MalzemeMaliyeti = dbContext.SIPARISLER
                        .Where(x => x.sip_stok_kod == item.rec_tuketim_kod && !x.sip_stok_kod.StartsWith("730.")) // rec_tuketim_kod kullanıldı
                        .OrderByDescending(x => x.sip_tarih)
                        .Select(x => new
                        {
                            item.rec_anakod,
                            x.sip_stok_kod,
                            x.sip_tarih,
                            TopFiyat = x.sip_b_fiyat * x.sip_miktar * x.sip_doviz_kuru,
                        })
                        .FirstOrDefault();


                    var MalzemeMaliyetiBirim = dbContext.SIPARISLER
                        .Where(x => x.sip_stok_kod == item.rec_tuketim_kod && !x.sip_stok_kod.StartsWith("730.")) // rec_tuketim_kod kullanıldı
                        .OrderByDescending(x => x.sip_tarih)
                        .Select(x => new
                        {
                            malzememaliyetibirim = x.sip_b_fiyat * x.sip_doviz_kuru,

                        })
                        .FirstOrDefault();

                    var İşçiMaliyetiTamamlanan123 = dbContext.URETIM_OPERASYON_HAREKETLERI
                        .Where(x => x.OpT_UrunKodu == item.rec_anakod)
                        .OrderByDescending(x => x.OpT_baslamatarihi)
                        .ThenBy(x => x.OpT_EvrakSatirNo)
                        .Select(x => new
                        {
                            x.OpT_IsEmriKodu,
                        })
                        .FirstOrDefault();

                    var İşçiMaliyetiPlanlanan = (İşçiMaliyetiTamamlanan123?.OpT_IsEmriKodu != null)
                        ? dbContext.URETIM_ROTA_PLANLARI
                            .Where(x => x.RtP_UrunKodu == item.rec_anakod
                                         && İşçiMaliyetiTamamlanan123.OpT_IsEmriKodu.Contains(x.RtP_IsEmriKodu))
                            .OrderByDescending(x => x.RtP_create_date)
                            .GroupBy(x => x.RtP_IsEmriKodu)
                            .Select(group => new
                            {
                                RtP_IsEmriKodu = group.Key,
                                TopFiyatIsci = group.Sum(x => x.RtP_PlanlananSure / (x.RtP_TamamlananMiktar != 0 ? x.RtP_TamamlananMiktar : 1)) / 60 * DKIsciMaliyet

                            })
                            .FirstOrDefault()
                        : null;




                    var İşçiMaliyetiTamamlanan = dbContext.URETIM_OPERASYON_HAREKETLERI
                        .Where(x => x.OpT_UrunKodu == item.rec_anakod && İşçiMaliyetiTamamlanan123.OpT_IsEmriKodu.Contains(x.OpT_IsEmriKodu))
                        .OrderByDescending(x => x.OpT_baslamatarihi)
                        .ThenBy(x => x.OpT_EvrakSatirNo)
                        .GroupBy(x => x.OpT_IsEmriKodu)
                          .Select(x => new
                          {
                              OpT_IsEmriKodu = x.Key,
                              TopFiyatIsciTam = x.Sum(x => x.OpT_TamamlananSure / x.OpT_TamamlananMiktar) / 60 * DKIsciMaliyet,

                          })
                        .FirstOrDefault();


                    var MalzemeSatisFiyatı = dbContext.SIPARISLER
                        .Where(x => x.sip_stok_kod == item.rec_anakod && x.sip_evrakno_seri == "")
                        .OrderByDescending(x => x.sip_tarih).Select(x => new
                        {
                            x.sip_stok_kod,
                            TopFiyatSatis = x.sip_b_fiyat * 1 * x.sip_doviz_kuru,

                        })
                        .FirstOrDefault();
                    int uzunluk = anakod.Length;

                    var MalzemeSatisFiyatı1 = dbContext.SIPARISLER
                         .Where(x => x.sip_stok_kod == item.rec_tuketim_kod && x.sip_evrakno_seri == "" && x.sip_aciklama.Substring(0, uzunluk) == anakod)
                         .OrderByDescending(x => x.sip_tarih)
                         .Select(x => new
                         {
                             x.sip_stok_kod,
                             TopFiyatSatis = x.sip_tutar * x.sip_doviz_kuru
                         })
                 .FirstOrDefault();

                    var MalzemeSatisFiyatıacikalama = dbContext.SIPARISLER
                        .Where(x => x.sip_stok_kod == item.rec_tuketim_kod
                                    && x.sip_evrakno_seri == ""
                                    && x.sip_aciklama.Substring(0, uzunluk) == anakod)
                        .OrderByDescending(x => x.sip_tarih)
                        .Select(x => new
                        {
                            sip_aciklama = x.sip_aciklama.Substring(uzunluk) // anakod dışındaki kısmı almak
                        })
                        .FirstOrDefault();

                    table.Rows.Add(
                          sayac++,
                          item.rec_anakod,
                          MalzemeMaliyeti?.sip_tarih.Date.ToString("yyyy/MM/dd"),
                          item.rec_tuketim_kod,
                          item.sto_isim,
                          item.rec_tuketim_miktar,
                          Math.Round(MalzemeMaliyetiBirim?.malzememaliyetibirim ?? 0, 2),
                          İşçiMaliyetiPlanlanan?.TopFiyatIsci,
                          İşçiMaliyetiTamamlanan?.TopFiyatIsciTam,
                          "1",
                          Math.Round(MalzemeSatisFiyatı?.TopFiyatSatis ?? 0, 2),
                          Math.Round(MalzemeSatisFiyatı1?.TopFiyatSatis ?? 0, 2),
                          MalzemeSatisFiyatıacikalama?.sip_aciklama
                      );
                }
                advancedDataGridView2.DataSource = table;
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            table.Columns.Clear();
            table.Rows.Clear();
            table.Clear();
            table.Columns.Add("SATIR_NO");
            table.Columns.Add("STOK KODU");
            table.Columns.Add("SİPARİŞ TARİHİ");
            table.Columns.Add("TÜKETİM KODU");
            table.Columns.Add("TÜKETİM KODU ADI");
            table.Columns.Add("TÜKETİM KODU KG MİKTAR");
            table.Columns.Add("BİRİM FİYATI");
            table.Columns.Add("PLANLANAN İŞÇİ MALİYETİ");
            table.Columns.Add("TAMAMLANAN İŞÇİ MALİYETİ");
            table.Columns.Add("TALEP MİKTARI");
            table.Columns.Add("SATINALMA FİYATI");
            table.Columns.Add("SATINALMA FİYATI TÜKETİM KOD");

            string anakod = textBox1.Text;
            int DKIsciMaliyet = Convert.ToInt32(textBox2.Text);
            double İşçiMaliyetiPlanlanan1 = 0;
            double İşçiMaliyetiTamamlanan1 = 0;
            var Talepler =
                    dbContext.URUN_RECETELERI.Where(x => x.rec_anakod.Substring(0, 13) == anakod)
                    .Select(x => new
                    {
                        x.rec_anakod,
                        x.rec_tuketim_kod,
                        x.rec_tuketim_miktar,
                        gereklimiktar = x.rec_tuketim_miktar * 1
                    }).Join(dbContext.STOKLAR,
                tuket => tuket.rec_tuketim_kod,
                stok => stok.sto_kod,
                    (tuket, stok) => new
                    {
                        tuket.rec_anakod,
                        tuket.rec_tuketim_kod,
                        tuket.gereklimiktar,
                        tuket.rec_tuketim_miktar,
                        stok.sto_isim
                    })
                .ToList();


            int sayac = 1;
            foreach (var item in Talepler)
            {
                var MalzemeMaliyeti = dbContext.SIPARISLER
                    .Where(x => x.sip_stok_kod == item.rec_tuketim_kod && item.rec_tuketim_kod.StartsWith("90.")) // rec_tuketim_kod kullanıldı
                    .OrderByDescending(x => x.sip_tarih)
                    .Select(x => new
                    {
                        item.rec_anakod,
                        x.sip_stok_kod,
                        x.sip_tarih,
                        TopFiyat = x.sip_b_fiyat * x.sip_miktar * x.sip_doviz_kuru,
                    })
                    .FirstOrDefault();


                var MalzemeMaliyetiBirim = dbContext.SIPARISLER
                    .Where(x => x.sip_stok_kod == item.rec_tuketim_kod && item.rec_tuketim_kod.StartsWith("90.")) // rec_tuketim_kod kullanıldı
                    .OrderByDescending(x => x.sip_tarih)
                    .Select(x => new
                    {
                        malzememaliyetibirim = x.sip_b_fiyat * x.sip_doviz_kuru,

                    })
                    .FirstOrDefault();

                var İşçiMaliyetiTamamlanan123 = dbContext.URETIM_OPERASYON_HAREKETLERI
                    .Where(x => x.OpT_UrunKodu == item.rec_anakod)
                    .OrderByDescending(x => x.OpT_baslamatarihi)
                    .ThenBy(x => x.OpT_EvrakSatirNo)
                    .Select(x => new
                    {
                        x.OpT_IsEmriKodu,
                    })
                    .FirstOrDefault();

                // İşçiMaliyetiTamamlanan123 null mı kontrolü
                if (İşçiMaliyetiTamamlanan123 != null)
                {
                    var İşçiMaliyetiPlanlanan = dbContext.URETIM_ROTA_PLANLARI
                        .Where(x => x.RtP_UrunKodu == item.rec_anakod
                                     && İşçiMaliyetiTamamlanan123.OpT_IsEmriKodu.Contains(x.RtP_IsEmriKodu))
                        .OrderByDescending(x => x.RtP_create_date)
                        .Select(x => new
                        {
                            TopFiyatIsci = (x.RtP_PlanlananSure / (x.RtP_TamamlananMiktar != 0 ? x.RtP_TamamlananMiktar : 1)) / 60 * DKIsciMaliyet
                        })
                        .FirstOrDefault(); // Eğer liste boşsa, varsayılan değer döner

                    // İşçiMaliyetiPlanlanan1'e atanması
                    İşçiMaliyetiPlanlanan1 = (İşçiMaliyetiPlanlanan?.TopFiyatIsci ?? 0); // Null kontrolü ile varsayılan değer

                    var İşçiMaliyetiTamamlanan = dbContext.URETIM_OPERASYON_HAREKETLERI
                        .Where(x => x.OpT_UrunKodu == item.rec_anakod
                                     && İşçiMaliyetiTamamlanan123.OpT_IsEmriKodu.Contains(x.OpT_IsEmriKodu))
                        .OrderByDescending(x => x.OpT_baslamatarihi)
                        .ThenBy(x => x.OpT_EvrakSatirNo)
                        .Select(x => new
                        {
                            TopFiyatIsciTam = (x.OpT_TamamlananSure / (x.OpT_TamamlananMiktar != 0 ? x.OpT_TamamlananMiktar : 1)) / 60 * DKIsciMaliyet
                        })
                        .FirstOrDefault(); // Eğer liste boşsa, varsayılan değer döner

                    // İşçiMaliyetiTamamlanan1'e atanması
                    İşçiMaliyetiTamamlanan1 = (İşçiMaliyetiTamamlanan?.TopFiyatIsciTam ?? 0); // Null kontrolü ile varsayılan değer
                }
                else
                {
                    // İşçiMaliyetiTamamlanan123 null ise, varsayılan değerler atama
                    İşçiMaliyetiPlanlanan1 = 0; // veya uygun bir varsayılan değer
                    İşçiMaliyetiTamamlanan1 = 0; // veya uygun bir varsayılan değer
                }




                var MalzemeSatisFiyatı = dbContext.SIPARISLER
                    .Where(x => x.sip_stok_kod == item.rec_anakod && x.sip_evrakno_seri == "")
                    .OrderByDescending(x => x.sip_tarih).Select(x => new
                    {
                        x.sip_stok_kod,
                        TopFiyatSatis = x.sip_b_fiyat * 1 * x.sip_doviz_kuru,

                    })
                    .FirstOrDefault();
                int uzunluk = anakod.Length;

                var MalzemeSatisFiyatı1 = dbContext.SIPARISLER
                     .Where(x => x.sip_stok_kod == item.rec_tuketim_kod && x.sip_evrakno_seri == "" && x.sip_aciklama.Substring(0, uzunluk) == anakod)
                     .OrderByDescending(x => x.sip_tarih)
                     .Select(x => new
                     {
                         x.sip_stok_kod,
                         TopFiyatSatis = x.sip_tutar * x.sip_doviz_kuru
                     })
             .FirstOrDefault();

                // DataTable'a yeni satır ekleme
                table.Rows.Add(
                    sayac++, // Satır numarası
                    item.rec_anakod, // Proje Kodu
                    MalzemeMaliyeti?.sip_tarih.Date.ToString("yyyy/MM/dd"), // Tarih
                    item.rec_tuketim_kod, // Tüketim Kodu
                    item.sto_isim, // Stok İsmi
                    item.rec_tuketim_miktar, // Talep Miktarı
                    Math.Round(MalzemeMaliyetiBirim?.malzememaliyetibirim ?? 0, 2), // Hammadde Maliyeti
                    Math.Round(İşçiMaliyetiPlanlanan1, 2), // Planlanan İşçi Maliyeti
                    Math.Round(İşçiMaliyetiTamamlanan1, 2), // Tamamlanan İşçi Maliyeti
                    "1", // Sabit Değer (örneğin, birim)
                    Math.Round(MalzemeSatisFiyatı?.TopFiyatSatis ?? 0, 2), // Satın Alma Fiyatı
                    Math.Round(MalzemeSatisFiyatı1?.TopFiyatSatis ?? 0, 2) // Fason Maliyeti
                );


            }
            advancedDataGridView2.DataSource = table;
        }

        //private void button3_Click(object sender, EventArgs e)
        //{
        //    table.Columns.Clear();
        //    table.Rows.Clear();
        //    table.Clear();
        //    table.Columns.Add("SATIR_NO");
        //    table.Columns.Add("SERİ");
        //    table.Columns.Add("SIRA");
        //    table.Columns.Add("SERİ-SIRA");
        //    table.Columns.Add("CARİ AD");
        //    table.Columns.Add("STOK KODU");
        //    table.Columns.Add("MİKTAR");
        //    table.Columns.Add("TESLİM MİKTARI");
        //    table.Columns.Add("TUTAR");

        //    DateTime filterDate = new DateTime(2023, 12, 31); // Tarih değerini DateTime olarak tanımlıyoruz

        //    var result = dbContext.SIPARISLER
        //        .Join(dbContext.CARI_HESAPLAR,
        //              sip => sip.sip_musteri_kod,
        //              cari => cari.cari_kod,
        //              (sip, cari) => new { sip, cari })
        //        .Where(x => x.sip.sip_projekodu.Contains("24/005") &&
        //                    x.sip.sip_create_date > filterDate && // DateTime kullanımı
        //                    !string.IsNullOrEmpty(x.sip.sip_evrakno_seri))
        //        .Select(x => new
        //        {
        //            x.sip.sip_evrakno_seri,
        //            x.sip.sip_evrakno_sira,
        //            evrakserisi = x.sip.sip_evrakno_seri + "-" + x.sip.sip_evrakno_sira,
        //            x.cari.cari_unvan1
        //        })
        //        .ToList();
        //    int sayac = 1;
        //    foreach (var item in result)
        //    {
        //        var result1 = dbContext.SIPARISLER.Where(x => x.sip_projekodu.Contains("24/005") &&
        //        x.sip_create_date > filterDate && x.sip_evrakno_seri == "" && x.sip_stok_sormerk == item.evrakserisi)
        //        .Select(x => new
        //        {
        //            x.sip_stok_kod,
        //            x.sip_miktar,
        //            x.sip_teslim_miktar,
        //            tutar = x.sip_tutar * x.sip_doviz_kuru,
        //        })
        //        .ToList();

        //        foreach (var item1 in result1)
        //        {
        //            // DataTable'a yeni satır ekleme
        //            table.Rows.Add(
        //                sayac++, // Satır numarası
        //                item.sip_evrakno_seri, // Proje Kodu
        //                item.sip_evrakno_sira,
        //                item.evrakserisi,
        //                item.cari_unvan1,
        //                item1.sip_stok_kod,
        //                item1.sip_miktar,
        //                item1.sip_teslim_miktar,
        //                Math.Round( item1.tutar,2)
        //            );
        //        }

        //    }
        //    advancedDataGridView2.DataSource = table;
        //}

    }
}
