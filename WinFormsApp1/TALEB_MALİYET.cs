using Microsoft.EntityFrameworkCore.Metadata.Internal;
using Microsoft.Office.Interop.Excel;
using System.Data;
using WinFormsApp1.Context;

namespace WinFormsApp1
{
    public partial class TALEB_MALİYET : Form
    {
        private MyDbContext dbContext;
        public TALEB_MALİYET()
        {
            dbContext = new MyDbContext();
            InitializeComponent();
        }
        System.Data.DataTable table = new System.Data.DataTable();
        private void button1_Click(object sender, EventArgs e)
        {
            table.Columns.Clear();
            table.Rows.Clear();
            table.Clear();
            table.Columns.Add("SATIR_NO");
            table.Columns.Add("PROJE KODU");
            table.Columns.Add("STOK KODU");
            table.Columns.Add("SİPARİŞ TARİHİ");
            table.Columns.Add("TÜKETİM KODU");
            table.Columns.Add("TÜKETİM KODU ADI");
            table.Columns.Add("TÜKETİM KODU KG MİKTAR");
            table.Columns.Add("(PRO*TUKET) MALZEME MİKTARI");
            table.Columns.Add("BİRİM FİYATI");
            table.Columns.Add("TOPLAM MALZEME MALİYETİ");
            table.Columns.Add("PLANLANAN İŞÇİ MALİYETİ");
            table.Columns.Add("TAMAMLANAN İŞÇİ MALİYETİ");
            table.Columns.Add("TALEP MİKTARI");
            table.Columns.Add("SATINALMA FİYATI");
            int EvranSiraNo = Convert.ToInt32(textBox1.Text);
            int DKIsciMaliyet = Convert.ToInt32(textBox2.Text);

            var Talepler = dbContext.PROFORMA_SIPARISLER
                .Where(x => x.pro_evrakno_sira == EvranSiraNo)
                .Select(x => new
                {
                    x.pro_evrakno_sira,
                    x.pro_stokkodu,
                    x.pro_miktar,
                    x.pro_projekodu
                })
                .Join(
                    dbContext.URUN_RECETELERI.Where(x => !x.rec_tuketim_kod.StartsWith("730.")),
                    talep => talep.pro_stokkodu,
                    recete => recete.rec_anakod,
                    (talep, recete) => new
                    {
                        talep.pro_evrakno_sira,
                        talep.pro_stokkodu,
                        talep.pro_projekodu,
                        talep.pro_miktar,
                        recete.rec_tuketim_kod,
                        recete.rec_tuketim_miktar,
                        gereklimiktar = recete.rec_tuketim_miktar * talep.pro_miktar
                    }
                ).Join(dbContext.STOKLAR,
                tuket => tuket.rec_tuketim_kod,
                stok => stok.sto_kod,
                    (tuket, stok) => new
                    {
                        tuket.pro_evrakno_sira,
                        tuket.pro_stokkodu,
                        tuket.pro_miktar,
                        tuket.rec_tuketim_kod,
                        tuket.gereklimiktar,
                        tuket.rec_tuketim_miktar,
                        tuket.pro_projekodu,
                        stok.sto_isim
                    })
                .ToList();


            int sayac = 1;
            foreach (var item in Talepler)
            {
                var MalzemeMaliyeti = dbContext.SIPARISLER
                    .Where(x => x.sip_stok_kod == item.rec_tuketim_kod) // rec_tuketim_kod kullanıldı
                    .OrderByDescending(x => x.sip_create_date)
                    .Select(x => new
                    {
                        item.pro_evrakno_sira,
                        x.sip_stok_kod,
                        x.sip_tarih,
                        TopFiyat = x.sip_b_fiyat * item.gereklimiktar,

                    })
                    .FirstOrDefault();

                var MalzemeMaliyetiBirim = dbContext.SIPARISLER
                    .Where(x => x.sip_stok_kod == item.rec_tuketim_kod) // rec_tuketim_kod kullanıldı
                    .OrderByDescending(x => x.sip_create_date)
                    .Select(x => new
                    {
                        x.sip_b_fiyat,

                    })
                    .FirstOrDefault();



                var İşçiMaliyetiPlanlanan = dbContext.URETIM_ROTA_PLANLARI
                        .Where(x => x.RtP_UrunKodu == item.pro_stokkodu)
                        .OrderByDescending(x => x.RtP_create_date).Select(x => new
                        {

                            TopFiyatIsci = (x.RtP_PlanlananSure / 60) * DKIsciMaliyet,

                        })
                        .FirstOrDefault();

                var İşçiMaliyetiTamamlanan = dbContext.URETIM_OPERASYON_HAREKETLERI
                    .Where(x => x.OpT_UrunKodu == item.pro_stokkodu)
                    .OrderByDescending(x => x.OpT_baslamatarihi).Select(x => new
                    {

                        TopFiyatIsciTam = (x.OpT_TamamlananSure / 60) * DKIsciMaliyet,

                    })
                    .FirstOrDefault();

                var MalzemeSatisFiyatı = dbContext.SIPARISLER
                    .Where(x => x.sip_stok_kod == item.pro_stokkodu)
                    .OrderByDescending(x => x.sip_create_date).Select(x => new
                    {
                        x.sip_stok_kod,
                        TopFiyatSatis = x.sip_b_fiyat * item.pro_miktar,

                    })
                    .FirstOrDefault();

                table.Rows.Add(
                      sayac++,
                      item.pro_projekodu,
                      item.pro_stokkodu,
                      MalzemeMaliyeti?.sip_tarih.Date.ToString("yyyy/MM/dd"),
                      item.rec_tuketim_kod,
                      item.sto_isim,
                      item.rec_tuketim_miktar,
                      Math.Round( item.gereklimiktar,0),
                      Math.Round(MalzemeMaliyetiBirim?.sip_b_fiyat ?? 0, 2),
                      Math.Round(MalzemeMaliyeti?.TopFiyat ?? 0, 2),
                      İşçiMaliyetiPlanlanan?.TopFiyatIsci,
                      İşçiMaliyetiTamamlanan?.TopFiyatIsciTam,
                      item.pro_miktar,
                      Math.Round(MalzemeSatisFiyatı?.TopFiyatSatis ?? 0, 2)
                  );


            }
            advancedDataGridView1.DataSource = table;
        }

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

        private void button2_Click(object sender, EventArgs e)
        {
            table.Columns.Clear();
            table.Rows.Clear();
            table.Clear();
            table.Columns.Add("SATIR_NO");
            table.Columns.Add("PROJE KODU");
            table.Columns.Add("STOK KODU");
            table.Columns.Add("TÜKETİM KODU");
            table.Columns.Add("TALEP MİKTARI");
            table.Columns.Add("TOPLAM MALZEME MALİYETİ");
            table.Columns.Add("PLANLANAN İŞÇİ MALİYETİ");
            table.Columns.Add("TAMAMLANAN İŞÇİ MALİYETİ");
            table.Columns.Add("SATINALMA FİYATI");
            table.Columns.Add("PLANLANAN MALİYETİ");
            table.Columns.Add("TAMAMLANAN MALİYETİ");
            int EvranSiraNo = Convert.ToInt32(textBox1.Text);
            int DKIsciMaliyet = Convert.ToInt32(textBox2.Text);

            var Talepler = dbContext.PROFORMA_SIPARISLER
                .Where(x => x.pro_evrakno_sira == EvranSiraNo)
                .Select(x => new
                {
                    x.pro_evrakno_sira,
                    x.pro_stokkodu,
                    x.pro_miktar,
                    x.pro_projekodu
                })
                .ToList();


            int sayac = 1;
            foreach (var item in Talepler)
            {
                var tuketimkod = dbContext.URUN_RECETELERI.Where(x => x.rec_anakod == item.pro_stokkodu && !x.rec_tuketim_kod.StartsWith("730."))
                 .Select(x => new
                 {
                     item.pro_evrakno_sira,
                     item.pro_stokkodu,
                     item.pro_miktar,
                     x.rec_tuketim_kod,
                     gereklimiktar = x.rec_tuketim_miktar * item.pro_miktar
                 }
             ).ToList();
                double fiyat = 0;
                double planlananMaliyet = 0;
                double TamamlananMaliyet = 0;
                foreach (var item2 in tuketimkod)
                {
                    var MalzemeMaliyeti = dbContext.SIPARISLER
                    .Where(x => x.sip_stok_kod == item2.rec_tuketim_kod) // rec_tuketim_kod kullanıldı
                    .OrderByDescending(x => x.sip_create_date)
                    .Select(x => new
                    {
                        TopFiyat = x.sip_b_fiyat * x.sip_doviz_kuru * item2.gereklimiktar,
                    })
                    .FirstOrDefault();
                    // Her bir TopFiyat değerini fiyat değişkenine ekle
                    fiyat += MalzemeMaliyeti?.TopFiyat ?? 0;
                }

                var İşçiMaliyetiPlanlanan = dbContext.URETIM_ROTA_PLANLARI
                    .Where(x => x.RtP_UrunKodu == item.pro_stokkodu)
                    .OrderByDescending(x => x.RtP_create_date)
                    .GroupBy(x => x.RtP_IsEmriKodu) // RtP_IsEmriKodu'na göre grupla
                    .Select(group => new
                    {
                        RtP_IsEmriKodu = group.Key, // Gruplanmış anahtar (IsEmriKodu)
                        TopFiyatIsci = group.Sum(x => x.RtP_PlanlananSure) / 60 * DKIsciMaliyet // Toplam işçi maliyeti hesapla
                    })
                    .FirstOrDefault(); // İlk grup için öğeyi seç

                var İşçiMaliyetiTamamlanan = dbContext.URETIM_OPERASYON_HAREKETLERI
                    .Where(x => x.OpT_UrunKodu == item.pro_stokkodu)
                    .OrderByDescending(x => x.OpT_baslamatarihi)
                    .GroupBy(x => x.OpT_IsEmriKodu)
                    .Select(x => new
                    {
                        OpT_IsEmriKodu = x.Key,
                        TopFiyatIsciTam = x.Sum(x => x.OpT_TamamlananSure) / 60 * DKIsciMaliyet,

                    })
                    .FirstOrDefault();

                var MalzemeSatisFiyatı = dbContext.SIPARISLER
                    .Where(x => x.sip_stok_kod == item.pro_stokkodu)
                    .OrderByDescending(x => x.sip_create_date).Select(x => new
                    {
                        x.sip_stok_kod,
                        TopFiyatSatis = x.sip_b_fiyat * item.pro_miktar,

                    })
                    .FirstOrDefault();
                planlananMaliyet = fiyat + (İşçiMaliyetiPlanlanan?.TopFiyatIsci ?? 0);
                TamamlananMaliyet = fiyat + (İşçiMaliyetiTamamlanan?.TopFiyatIsciTam ?? 0);


                table.Rows.Add(
                  sayac++,
                  item.pro_projekodu,
                  item.pro_stokkodu,
                  string.Join(" ; ", tuketimkod.Select(tk => tk.rec_tuketim_kod)),
                  item.pro_miktar,
                  Math.Round(fiyat, 2),
                  İşçiMaliyetiPlanlanan?.TopFiyatIsci,
                  İşçiMaliyetiTamamlanan?.TopFiyatIsciTam,
                  Math.Round(MalzemeSatisFiyatı?.TopFiyatSatis ?? 0, 2),
                  Math.Round(planlananMaliyet, 2),
                  Math.Round(TamamlananMaliyet, 2)
              );


            }
            advancedDataGridView1.DataSource = table;
        }

        private void TALEB_MALİYET_Load(object sender, EventArgs e)
        {
            toolStripLabel1.ToolTipText =

                "1.Evrak Sıra No Yaz\n" +

                "2.İşçilerin Dakikalık Ücretini Yaz\n" +

                "3.Listele";
        }
    }
}
