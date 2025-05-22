using DevExpress.Xpo.DB;
using Microsoft.EntityFrameworkCore;
using SİPARİŞ_GİRİŞ.Context;
using SİPARİŞ_GİRİŞ.Tables;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WinFormsApp1.Context;
using WinFormsApp1.Tables;
using Zuby.ADGV;

namespace SİPARİŞ_GİRİŞ
{
    public partial class AĞAC_TERMİN : Form
    {
        private MyDbContext dbContext;
        private TLBContext TLBdbContext;

        // Listeyi sınıf düzeyinde tanımla

        public AĞAC_TERMİN()
        {
            dbContext = new MyDbContext();
            TLBdbContext = new TLBContext();
            InitializeComponent();
        }
        // DataTable oluştur
        DataTable dt = new DataTable();

        private void button1_Click(object sender, EventArgs e)
        {
            // TextBox'tan LRF değerini al
            int lrf = Convert.ToInt32(textBox1.Text);
            // Veriyi önce çek, sonra işle
            var eşleştir = dbContext.zz_bom
                .Where(x => x.dsref == lrf)
                .Select(x => new
                {
                    x.item_partnumber,
                    x.item_sirano,
                    x.title,
                    x.rotakod1,
                    x.rotakod2,
                    x.rotakod3,
                    x.rotakod4,
                    x.rotakod5,
                    x.rotakod6,
                    x.rotakod7,
                    x.rotakod8,
                    x.rotakod9,
                    x.rotakod10,
                    x.material,
                    x.hammadde_erp_kodu,
                    fason = x.fason ?? "HAYIR"  // NULL ise "HAYIR" olarak ata
                })
                .AsEnumerable() // SQL'den çekip bellekte işleme geçiyoruz
                .Select(x => new
                {
                    x.item_partnumber,
                    x.item_sirano,
                    x.title,
                    x.fason,
                    x.rotakod1,
                    x.rotakod2,
                    x.rotakod3,
                    x.rotakod4,
                    x.rotakod5,
                    x.rotakod6,
                    x.rotakod7,
                    x.rotakod8,
                    x.rotakod9,
                    x.rotakod10,
                    x.material,
                    x.hammadde_erp_kodu,
                    Level = x.item_sirano == "0" ? 0 : x.item_sirano.Count(c => c == '.') + 1 // Level hesaplama
                })
                .ToList();


            // Maksimum Level'ı belirle
            int maxLevel = eşleştir.Max(x => x.Level);

            // Level ve Sure sütunlarını ekle
            for (int i = 0; i <= maxLevel; i++)
            {
                dt.Columns.Add($"LEVEL {i}", typeof(string));
                dt.Columns.Add($"SURE {i}", typeof(int)); // Sipariş süresi için sütun
            }

            // item_sirano'su "0" olan öğeyi bul
            var itemZero = eşleştir.FirstOrDefault(x => x.item_sirano == "0");

            // Her satır için yeni DataRow oluştur
            foreach (var item in eşleştir)
            {

                DataRow row = dt.NewRow();

                // Level 0 için item_partnumber'ı yerleştir
                if (item.item_sirano == "0")
                {
                    row["Level 0"] = item.item_partnumber;
                }

                // Level bazlı sütunlara item_partnumber değerlerini yerleştir
                if (item.item_sirano != "0")
                {
                    string[] levels = item.item_sirano.Split('.');

                    if (levels.Length > 0) row["Level 1"] = levels[0];
                    if (levels.Length > 1) row["Level 2"] = string.Join(".", levels.Take(2));
                    if (levels.Length > 2) row["Level 3"] = string.Join(".", levels.Take(3));
                    if (levels.Length > 3) row["Level 4"] = string.Join(".", levels.Take(4));
                    if (levels.Length > 4) row["Level 5"] = string.Join(".", levels.Take(5));
                }

                dt.Rows.Add(row);
            }

            // Eğer item_sirano'su "0" olan bir öğe varsa, tüm Level 0 satırlarını güncelle
            if (itemZero != null)
            {
                foreach (DataRow dataRow in dt.Rows)
                {
                    dataRow["Level 0"] = itemZero.item_partnumber;
                }
            }

            // DataTable'daki her hücre için item_sirano değerini güncelle
            foreach (DataRow dataRow in dt.Rows)
            {
                for (int level = 0; level <= maxLevel; level++)
                {
                    string item_sirano = dataRow[$"Level {level}"]?.ToString();

                    // item_sirano değerine karşılık gelen item_partnumber'ı bul
                    var eşleşenItem = eşleştir.FirstOrDefault(x => x.item_sirano == item_sirano);

                    // Eğer eşleşme varsa ve Level 0 değilse, şartlara göre uygun değeri yaz
                    if (eşleşenItem != null && level != 0) // Level 0 hariç diğerlerini güncelle
                    {
                        string yeniDeger; // Varsayılan olarak item_partnumber kullan

                        bool fasonHayir = eşleşenItem.fason?.Trim() == "HAYIR"; // Trim ile boşlukları temizle
                        bool hammaddeDolu = !string.IsNullOrWhiteSpace(eşleşenItem.hammadde_erp_kodu); // Null veya boş mu kontrol et
                        bool baslangic90 = eşleşenItem.hammadde_erp_kodu?.Trim().StartsWith("90.") ?? false; // Trim ile boşlukları temizle

                        if (fasonHayir && hammaddeDolu && !baslangic90)
                        {
                            yeniDeger = eşleşenItem.hammadde_erp_kodu; // Şarta uyuyorsa hammadde_erp_kodu'nu yaz
                        }
                        else
                        {
                            yeniDeger = eşleşenItem.item_partnumber;
                        }

                        // Güncellenen değeri hücreye yaz
                        dataRow[$"Level {level}"] = yeniDeger;
                    }
                }
            }

            // **STOKLAR tablosundan sto_siparis_sure değerlerini çekip toplama işlemi**
            foreach (DataRow dataRow in dt.Rows)
            {
                for (int level = 0; level <= maxLevel; level++)
                {
                    string itemPartNumber = dataRow[$"Level {level}"]?.ToString();
                    if (!string.IsNullOrEmpty(itemPartNumber))
                    {
                        // itemPartNumber veya uygun hammadde_erp_kodu ile eşleşen veriyi bul
                        var eşleşenItem = eşleştir.FirstOrDefault(x =>
                            x.item_partnumber == itemPartNumber || x.hammadde_erp_kodu == itemPartNumber &&
                            (!string.IsNullOrEmpty(x.hammadde_erp_kodu) && !x.hammadde_erp_kodu.StartsWith("90."))
                        );


                        if (eşleşenItem != null)
                        {
                            string stokKodu = itemPartNumber.Length >= 13 ? itemPartNumber.Substring(0, 13) : itemPartNumber;

                            // Eğer item_partnumber "01", "02" veya "99" ile başlıyorsa
                            if ((eşleşenItem.material == "COMPONENT" || eşleşenItem.material == "TEKNIKRESIM")&&eşleşenItem.fason=="HAYIR")
                            {
                                // İlk 13 hanesi stokKodu ile eşleşen tüm stokları getir
                                var stoklar = dbContext.STOKLAR
                                    .Where(s => s.sto_kod == itemPartNumber)
                                    .ToList();

                                int toplamSure = 0;

                                // Eşleşen tüm stokların sto_siparis_sure değerlerini topla
                                foreach (var stok in stoklar)
                                {

                                    toplamSure = Convert.ToInt32(stok.sto_siparis_sure);

                                }

                                // Sure sütununa yaz
                                dataRow[$"Sure {level}"] = toplamSure;


                            }
                            else // Eğer 01, 02, 99 ile başlamıyorsa
                            {
                                var stoklar = dbContext.STOKLAR
                                          .Where(s => s.sto_kod.StartsWith(stokKodu))
                                          .ToList();

                                int toplamSure = 0;

                                // Rotalardan gelen son 3 haneyi al ve eşleşen sto_siparis_sure değerlerinden sadece ilkini topla
                                var rotaKodlari = new List<string>
                                        {
                                            eşleşenItem.rotakod1, eşleşenItem.rotakod2, eşleşenItem.rotakod3,
                                            eşleşenItem.rotakod4, eşleşenItem.rotakod5, eşleşenItem.rotakod6,
                                            eşleşenItem.rotakod7, eşleşenItem.rotakod8, eşleşenItem.rotakod9,
                                            eşleşenItem.rotakod10
                                        }.Where(kod => !string.IsNullOrEmpty(kod)) // Null veya boş olmayanları filtrele
                                                     .Select(kod => kod.Length >= 3 ? kod[^3..] : kod) // Son 3 haneyi al
                                                     .ToList();

                                foreach (var rotaKod in rotaKodlari)
                                {
                                    var eşleşenStok = stoklar.FirstOrDefault(stok => stok.sto_kod.EndsWith(rotaKod));
                                    if (eşleşenStok != null)
                                    {
                                        toplamSure += eşleşenStok.sto_siparis_sure;
                                    }
                                }

                                // Sure sütununa yaz
                                dataRow[$"Sure {level}"] = toplamSure;



                            }
                        }
                    }
                }
            }





            // Yeni bir sütun ekle: "Toplam Sure"
            dt.Columns.Add("TOPLAM SURE", typeof(int));

            foreach (DataRow dataRow in dt.Rows)
            {
                int toplamSure = 0;

                // Sure sütunlarındaki değerleri topla
                for (int level = 0; level <= maxLevel; level++)
                {
                    var levelColumnName = $"Level {level}";  // Level sütunlarının adı
                    var sureColumnName = $"Sure {level}";    // Sure sütununun adı

                    // Eğer ilgili sütunlar yoksa devam et
                    if (!dt.Columns.Contains(levelColumnName) || !dt.Columns.Contains(sureColumnName)) continue;

                    var levelValue = dataRow[levelColumnName]; // Level hücresindeki değer
                    var sureValue = dataRow[sureColumnName];   // Sure hücresindeki değer

                    // Eğer Sure sütunu dolu değilse devam et
                    if (sureValue == DBNull.Value || sureValue == null || string.IsNullOrWhiteSpace(sureValue.ToString())) continue;

                    int sure = Convert.ToInt32(sureValue);

                    // Eğer Level doluysa ve eşleşen bir item_partnumber varsa
                    if (levelValue != DBNull.Value && levelValue != null && !string.IsNullOrWhiteSpace(levelValue.ToString()))
                    {
                        var matchingItem = eşleştir.Where(x => x.material == "COMPONENT" || x.material == "TEKNIKRESIM"||x.fason=="EVET").FirstOrDefault(x => x.item_partnumber == levelValue.ToString());

                        // Eğer eşleşme varsa, süreyi toplamaya dahil etme
                        if (matchingItem != null)
                        {
                            continue;
                        }
                    }

                    // Eşleşme yoksa süreyi topla
                    toplamSure += sure;
                }

                // Hesaplanan toplam süreyi yeni sütuna yaz
                dataRow["TOPLAM SURE"] = toplamSure;
            }



            // Güncellenmiş DataTable'i AdvancedDataGridView'e bağla
            advancedDataGridView1.DataSource = dt;



        }

        private void button2_Click(object sender, EventArgs e)
        {
            DateTime makinateslimTarihi = dateTimePicker1.Value.Date;

            // "SİPARİŞ TARİHİ" adında yeni bir DateTime sütunu ekle
            if (!dt.Columns.Contains("TESLİM TARİHİ"))
            {
                dt.Columns.Add("TESLİM TARİHİ", typeof(DateTime));
            }

            // Her satır için "SİPARİŞ TARİHİ"ni hesapla
            foreach (DataRow dataRow in dt.Rows)
            {
                // "SİPARİŞ ZAMANI" değerini al (gün cinsinden)
                int sipSure = dataRow.Field<int?>("Toplam Sure") ?? 0; // NULL kontrolü

                // Gün aralıklarına göre küçük olana yuvarlama (haftalık - 7 gün periyotları)
                int baslangicGun = (sipSure / 7) * 7;

                // Sipariş tarihini makinateslimTarihi'nden geriye hesapla
                DateTime siparisTarihi = makinateslimTarihi.AddDays(-baslangicGun);

                // Hesaplanan sipariş tarihini yeni sütuna yaz
                dataRow["TESLİM TARİHİ"] = siparisTarihi;
            }




            // Güncellenmiş DataTable'i AdvancedDataGridView'e bağla
            advancedDataGridView1.DataSource = dt;
        }


        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            // DataGridView'deki verileri bir DataTable'a kopyalayın
            DataTable dt = new DataTable();

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
                    dataRow[cell.ColumnIndex] = cell.Value;
                }
                dt.Rows.Add(dataRow);
            }

            // Excel uygulamasını başlatın
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;

            // Yeni bir Excel çalışma kitabı oluşturun
            Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];

            // DataTable'ı Excel çalışma sayfasına aktarın (tablo başlıklarını da dahil etmek için)
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
                    worksheet.Cells[rowIndex, j + 1] = dt.Rows[i][j].ToString();
                }
            }
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            // DataTable'ı tamamen temizle
            dt.Dispose(); // Önce DataTable'ı serbest bırak
            dt = new DataTable(); // Yeni bir DataTable oluştur

            // AdvancedDataGridView'i temizle
            advancedDataGridView1.DataSource = null;  // Bağlantıyı kaldır
            advancedDataGridView1.Rows.Clear(); // Satırları temizle
            advancedDataGridView1.Columns.Clear(); // Sütunları temizle
            advancedDataGridView1.Refresh(); // Güncellemeleri uygula
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            // Yeni bir form oluştur
            Form treeViewForm = new Form
            {
                ClientSize = new System.Drawing.Size(1920, 1080),
                Text = "TreeView Form",
                FormBorderStyle = FormBorderStyle.Sizable
            };

            // TreeView kontrolünü oluştur
            TreeView treeView = new TreeView
            {
                Dock = DockStyle.Fill
            };

            // AdvancedDataGridView'den değerleri al
            var rows = advancedDataGridView1.Rows.Cast<DataGridViewRow>()
                .Where(row => row.Cells["Level 0"].Value != null).ToList();

            // Hangi level'ların mevcut olduğunu kontrol et
            bool hasLevel1 = advancedDataGridView1.Columns.Contains("Level 1");
            bool hasLevel2 = advancedDataGridView1.Columns.Contains("Level 2");
            bool hasLevel3 = advancedDataGridView1.Columns.Contains("Level 3");
            bool hasLevel4 = advancedDataGridView1.Columns.Contains("Level 4");
            bool hasLevel5 = advancedDataGridView1.Columns.Contains("Level 5");

            // Hangi sure sütunlarının mevcut olduğunu kontrol et
            bool hasSure1 = advancedDataGridView1.Columns.Contains("Sure 1");
            bool hasSure2 = advancedDataGridView1.Columns.Contains("Sure 2");
            bool hasSure3 = advancedDataGridView1.Columns.Contains("Sure 3");
            bool hasSure4 = advancedDataGridView1.Columns.Contains("Sure 4");
            bool hasSure5 = advancedDataGridView1.Columns.Contains("Sure 5");

            // Level 0 düğümlerini oluştur
            var level0Groups = rows.GroupBy(row => row.Cells["Level 0"].Value.ToString());

            foreach (var level0Group in level0Groups)
            {
                // Level 0 düğümünü ekle (Sure eklenmez)
                TreeNode level0Node = new TreeNode(level0Group.Key);
                treeView.Nodes.Add(level0Node);

                // Level 1 değerlerini kontrol et ve ekle
                if (hasLevel1)
                {
                    var level1Groups = level0Group.GroupBy(row => row.Cells["Level 1"].Value?.ToString()).Where(g => !string.IsNullOrEmpty(g.Key));

                    foreach (var level1Group in level1Groups)
                    {
                        string sureValue = GetSureValue(level1Group.FirstOrDefault(), 1);
                        TreeNode level1Node = new TreeNode($"{level1Group.Key} - {sureValue}");
                        level0Node.Nodes.Add(level1Node);

                        if (hasLevel2)
                        {
                            var level2Groups = level1Group.GroupBy(row => row.Cells["Level 2"].Value?.ToString()).Where(g => !string.IsNullOrEmpty(g.Key));

                            foreach (var level2Group in level2Groups)
                            {
                                sureValue = GetSureValue(level2Group.FirstOrDefault(), 2);
                                TreeNode level2Node = new TreeNode($"{level2Group.Key} - {sureValue}");
                                level1Node.Nodes.Add(level2Node);

                                if (hasLevel3)
                                {
                                    var level3Groups = level2Group.GroupBy(row => row.Cells["Level 3"].Value?.ToString()).Where(g => !string.IsNullOrEmpty(g.Key));

                                    foreach (var level3Group in level3Groups)
                                    {
                                        sureValue = GetSureValue(level3Group.FirstOrDefault(), 3);
                                        TreeNode level3Node = new TreeNode($"{level3Group.Key} - {sureValue}");
                                        level2Node.Nodes.Add(level3Node);

                                        if (hasLevel4)
                                        {
                                            var level4Groups = level3Group.GroupBy(row => row.Cells["Level 4"].Value?.ToString()).Where(g => !string.IsNullOrEmpty(g.Key));

                                            foreach (var level4Group in level4Groups)
                                            {
                                                sureValue = GetSureValue(level4Group.FirstOrDefault(), 4);
                                                TreeNode level4Node = new TreeNode($"{level4Group.Key} - {sureValue}");
                                                level3Node.Nodes.Add(level4Node);

                                                if (hasLevel5)
                                                {
                                                    var level5Groups = level4Group.GroupBy(row => row.Cells["Level 5"].Value?.ToString()).Where(g => !string.IsNullOrEmpty(g.Key));

                                                    foreach (var level5Group in level5Groups)
                                                    {
                                                        sureValue = GetSureValue(level5Group.FirstOrDefault(), 5);
                                                        TreeNode level5Node = new TreeNode($"{level5Group.Key} - {sureValue}");
                                                        level4Node.Nodes.Add(level5Node);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // TreeView'ı forma ekle
            treeViewForm.Controls.Add(treeView);

            // Tüm düğümleri aç
            treeView.ExpandAll();

            // Formu aç
            treeViewForm.Show();

            // Belirli bir level için ilgili sure değerini döndüren fonksiyon
            string GetSureValue(DataGridViewRow row, int level)
            {
                if (row == null) return "0";

                switch (level)
                {
                    case 1: return row.Cells["Sure 1"]?.Value?.ToString() ?? "0";
                    case 2: return row.Cells["Sure 2"]?.Value?.ToString() ?? "0";
                    case 3: return row.Cells["Sure 3"]?.Value?.ToString() ?? "0";
                    case 4: return row.Cells["Sure 4"]?.Value?.ToString() ?? "0";
                    case 5: return row.Cells["Sure 5"]?.Value?.ToString() ?? "0";
                    default: return "0";
                }
            }

        }
        // Satın alma talepleri için bir liste tanımlayın
        List<SATINALMA_TALEPLERI1> satınalmaList = new List<SATINALMA_TALEPLERI1>();
        private void button4_Click(object sender, EventArgs e)
        {
            // TextBox'tan LRF değerini al
            int lrf = Convert.ToInt32(textBox1.Text);
            // Veriyi önce çek, sonra işle
            var eşleştir = dbContext.zz_bom
                .Where(x => x.dsref == lrf)
                .Select(x => new
                {
                    x.item_partnumber,
                    x.item_sirano,
                    x.title,
                    x.rotaitemkod1,
                    x.material,
                    x.hammadde_erp_kodu,
                    fason = x.fason ?? "HAYIR"  // NULL ise "HAYIR" olarak ata
                })
                .AsEnumerable() // SQL'den çekip bellekte işleme geçiyoruz
                .Select(x => new
                {
                    x.item_partnumber,
                    x.item_sirano,
                    x.title,
                    x.fason,
                    x.rotaitemkod1,
                    x.material,
                    x.hammadde_erp_kodu,
                    Level = x.item_sirano == "0" ? 0 : x.item_sirano.Count(c => c == '.') + 1 // Level hesaplama
                })
                .ToList();
            // Mevcut Level sütunlarını bul
            var mevcutLevelSutunlari = advancedDataGridView1.Columns
                .Cast<DataGridViewColumn>()
                .Where(c => c.Name.StartsWith("LEVEL")) // LEVEL ile başlayan sütunları bul
                .Select(c => new
                {
                    Level = int.TryParse(c.Name.Replace("LEVEL ", ""), out int level) ? level : -1, // LEVEL x’ten x değerini al
                    ColumnName = c.Name
                })
                .Where(x => x.Level >= 0) // Geçerli Level’leri al
                .OrderByDescending(x => x.Level) // En yüksek Level’dan başlayarak sırala
                .ToList();

            // Eğer hiç LEVEL sütunu yoksa işlemi durdur
            if (!mevcutLevelSutunlari.Any()) return;

            // Teslim tarihlerini saklayacağımız sözlük
            var teslimTarihleri = new Dictionary<string, string>();

            // DataGridView'deki verileri al
            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                // Teslim tarihini al
                string teslimTarihi = row.Cells["TESLİM TARİHİ"].Value?.ToString();

                // Tüm mevcut Level sütunlarını kontrol et (En yüksek Level’dan başlayarak)
                for (int i = 0; i < mevcutLevelSutunlari.Count; i++)
                {
                    var currentLevel = mevcutLevelSutunlari[i]; // Şu anki Level sütunu
                    var nextLevel = (i + 1 < mevcutLevelSutunlari.Count) ? mevcutLevelSutunlari[i + 1] : null; // Bir alt Level varsa onu al

                    string hucreKodu = row.Cells[currentLevel.ColumnName].Value?.ToString(); // O hücredeki kod
                    string sonrakiLevel = nextLevel != null ? row.Cells[nextLevel.ColumnName].Value?.ToString() : null; // Bir alt Level hücresi

                    if (!string.IsNullOrEmpty(hucreKodu)) // Hücre doluysa
                    {
                        // Eğer teslim tarihi yoksa, önceki seviyeden al
                        if (!teslimTarihleri.ContainsKey(hucreKodu))
                        {
                            teslimTarihleri[hucreKodu] = teslimTarihi;
                        }
                    }
                }
            }

            // Şimdi en düşük Level'dan (Level 0) başlayarak eksik teslim tarihlerini tamamlayalım
            foreach (var level in mevcutLevelSutunlari.OrderBy(x => x.Level)) // Level 0'dan itibaren sırala
            {
                foreach (DataGridViewRow row in advancedDataGridView1.Rows)
                {
                    string hucreKodu = row.Cells[level.ColumnName].Value?.ToString(); // O Level'daki kod

                    if (!string.IsNullOrEmpty(hucreKodu))
                    {
                        // Eğer bu kod için bir teslim tarihi atanmışsa, alt Level'lara da ekle
                        if (teslimTarihleri.ContainsKey(hucreKodu))
                        {
                            string teslimTarihi = teslimTarihleri[hucreKodu];

                            // Alt Level'daki tüm kodları kontrol et ve eksik tarihleri tamamla
                            foreach (var altLevel in mevcutLevelSutunlari.Where(l => l.Level > level.Level))
                            {
                                string altKod = row.Cells[altLevel.ColumnName].Value?.ToString();
                                if (!string.IsNullOrEmpty(altKod) && !teslimTarihleri.ContainsKey(altKod))
                                {
                                    teslimTarihleri[altKod] = teslimTarihi;
                                }
                            }
                        }
                    }
                }
            }

            // Eşleştir listesi ile teslim tarihlerini birleştir
            var yeniListe = eşleştir
                .Select(e => new
                {
                    e.item_partnumber,
                    e.item_sirano,
                    e.title,
                    e.fason,
                    e.rotaitemkod1,
                    e.material,
                    e.hammadde_erp_kodu,
                    TeslimTarihi = teslimTarihleri.ContainsKey(e.item_partnumber) ? teslimTarihleri[e.item_partnumber] : null
                })
                .ToList();

            // Yeni eşleşen listeyi oluştur
            var eslesenListe = satınalmaList
                .Select(satinalma =>
                {
                    // Öncelikli olarak hammadde koduyla eşleşme
                    var eslesme = yeniListe.FirstOrDefault(yeni =>
                        (!string.IsNullOrEmpty(yeni.hammadde_erp_kodu) && yeni.fason != "EVET" && yeni.hammadde_erp_kodu == satinalma.stl_Stok_kodu) ||
                        (string.IsNullOrEmpty(yeni.hammadde_erp_kodu) || yeni.fason == "EVET") && yeni.item_partnumber == satinalma.stl_Stok_kodu
                    );

                    // Eğer hammadde ERP kodu ile eşleştiyse, onu al; aksi halde item_partnumber'ı al
                    string eslesenKod = (!string.IsNullOrEmpty(eslesme?.hammadde_erp_kodu) && eslesme.fason != "EVET")
                        ? eslesme.hammadde_erp_kodu
                        : eslesme?.item_partnumber;

                    return new
                    {
                        Eslestirilen_Kod = eslesenKod, // Hammadde ERP kodu veya item_partnumber buraya gelecek
                        fason = eslesme?.fason,
                        sırano = satinalma.stl_evrak_sira,
                        Eslestirilen_TeslimTarihi = eslesme?.TeslimTarihi
                    };
                })
                .ToList();

            foreach (var item in eslesenListe)
            {
                DateTime? yeniTeslimTarihi = null;

                // String'i DateTime'a çevirme
                if (!string.IsNullOrWhiteSpace(item.Eslestirilen_TeslimTarihi) &&
                    DateTime.TryParse(item.Eslestirilen_TeslimTarihi, out DateTime parsedDate))
                {
                    yeniTeslimTarihi = parsedDate; // Dönüştürülmüş tarihi atayın
                }

                // Eğer yeni teslim tarihi bugünden önceyse, dateTimePicker1.Value.Date - 7 gün kullan
                if (!yeniTeslimTarihi.HasValue || yeniTeslimTarihi.Value.Date < DateTime.Today)
                {
                    yeniTeslimTarihi =DateTime.Now.Date;
                }

                // Güncelleme yapmadan önce değerleri kontrol et
                Console.WriteLine($"Güncellenecek Kayıt: Evrak Sıra: {item.sırano}, Stok Kodu: {item.Eslestirilen_Kod}, Yeni Teslim Tarihi: {yeniTeslimTarihi}");

                // SQL sorgusu ile güncelleme yapın
                string sql = "UPDATE SATINALMA_TALEPLERI SET stl_teslim_tarihi = {0} WHERE stl_evrak_sira = {1} AND stl_Stok_kodu = {2}";

                // `null` ise `dateTimePicker1.Value.Date - 7 gün` kullan
                var parametre1 = yeniTeslimTarihi ?? dateTimePicker1.Value.Date.AddDays(-7);

                var result = TLBdbContext.Database.ExecuteSqlRaw(sql, parametre1, item.sırano, item.Eslestirilen_Kod);
            }

            MessageBox.Show("Başarıyla Güncellendi...");


        }
        private void button3_Click(object sender, EventArgs e)
        {
            // Kullanıcıdan alınan sıra numarasını tamsayıya dönüştürün
            int sirano = Convert.ToInt32(textBox3.Text);

            // Veritabanından satın alma taleplerini alın
            var satınalma = TLBdbContext.SATINALMA_TALEPLERI1.Where(x => x.stl_evrak_sira == sirano)
                .ToList();

            // Satın alma taleplerini satınalmaList'e ekleyin
            satınalmaList.AddRange(satınalma);

            MessageBox.Show("Eklendi...");
        }
        // Public olarak liste tanımla
         List<string> PublicVeriListesi = new List<string>();
        private void button5_Click(object sender, EventArgs e)
        {
            string isemrikodu = textBox2.Text;
            DateTime ilktarih = dateTimePicker2.Value.Date;
            DateTime sontarih = dateTimePicker3.Value.Date;
            int Durum = Convert.ToInt32(comboBox1.Text.ToString());

            var sorgu = from isEmir in dbContext.ISEMIRLERI
                        join uretim in dbContext.URETIM_MALZEME_PLANLAMA on isEmir.is_Kod equals uretim.upl_isemri
                        where isEmir.is_Kod.Substring(0, 8) == isemrikodu &&
                              isEmir.is_create_date.Date >= ilktarih &&
                              isEmir.is_create_date.Date <= sontarih &&
                              isEmir.is_EmriDurumu == Durum && // Durum kontrolü
                              dbContext.ISEMIRLERI_USER.Any(u => u.Record_uid == isEmir.is_Guid && u.is_emri_tipi != null) && // Null kontrolü
                              uretim.upl_satirno != 0
                        orderby uretim.upl_kodu ascending
                        select new
                        {
                            PROJE_KOD = isEmir.is_ProjeKodu,
                            İŞ_EMRİ_KODU = isEmir.is_Kod,
                            ÜRETİLECEK_ÜRÜN_KODU = uretim.upl_urstokkod,
                        };




            // Sorgu sonucundaki verileri liste olarak alın
            var veriListesi = sorgu.ToList();
            PublicVeriListesi = sorgu
                .Select(x => $"{x.PROJE_KOD} - {x.İŞ_EMRİ_KODU} - {x.ÜRETİLECEK_ÜRÜN_KODU}")
                .ToList();

            MessageBox.Show("Başarıyla Eklendi...");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            // TextBox'tan LRF değerini al
            int lrf = Convert.ToInt32(textBox1.Text);

            // Veriyi önce çek, sonra işle
            var eşleştir = dbContext.zz_bom
                .Where(x => x.dsref == lrf)
                .Select(x => new
                {
                    x.item_partnumber,
                    x.item_sirano,
                    x.title,
                    x.rotaitemkod1,
                    x.rotaitemkod2,
                    x.rotaitemkod3,
                    x.rotaitemkod4,
                    x.rotaitemkod5,
                    x.rotaitemkod6,
                    x.rotaitemkod7,
                    x.rotaitemkod8,
                    x.rotaitemkod9,
                    x.rotaitemkod10,
                    x.material,
                    x.hammadde_erp_kodu,
                    fason = x.fason ?? "HAYIR"  // NULL ise "HAYIR" olarak ata
                })
                .AsEnumerable() // SQL'den çekip bellekte işleme geçiyoruz
                .Select(x => new
                {
                    x.item_partnumber,
                    x.item_sirano,
                    x.title,
                    x.fason,
                    x.rotaitemkod1,
                    x.rotaitemkod2,
                    x.rotaitemkod3,
                    x.rotaitemkod4,
                    x.rotaitemkod5,
                    x.rotaitemkod6,
                    x.rotaitemkod7,
                    x.rotaitemkod8,
                    x.rotaitemkod9,
                    x.rotaitemkod10,
                    x.material,
                    x.hammadde_erp_kodu,
                    Level = x.item_sirano == "0" ? 0 : x.item_sirano.Count(c => c == '.') + 1 // Level hesaplama
                })
                .ToList();

            // Yeni liste oluştur
            var transformedList = new List<dynamic>();

            foreach (var item in eşleştir)
            {
                // Rota kodlarını ekle
                var rotaKodlari = new List<string>
                    {
                        item.rotaitemkod1, item.rotaitemkod2, item.rotaitemkod3, item.rotaitemkod4,
                        item.rotaitemkod5, item.rotaitemkod6, item.rotaitemkod7, item.rotaitemkod8,
                        item.rotaitemkod9, item.rotaitemkod10
                    };

                foreach (var rota in rotaKodlari.Where(r => !string.IsNullOrEmpty(r)))
                {
                    // STOKLAR tablosundan sto_siparis_sure değerini al
                    var stokBilgi = dbContext.STOKLAR.FirstOrDefault(s => s.sto_kod == rota);

                    transformedList.Add(new
                    {
                        item_partnumber = rota,
                        item_sirano = "", // Boş bırakabiliriz
                        title = "",
                        fason = "",
                        material = "",
                        hammadde_erp_kodu = "",
                        sto_siparis_sure = stokBilgi.sto_siparis_sure, // Eşleşmezse boş bırak
                        Level = item.Level + 1
                    });
                }
            }







            // Mevcut Level sütunlarını bul
            var mevcutLevelSutunlari = advancedDataGridView1.Columns
                .Cast<DataGridViewColumn>()
                .Where(c => c.Name.StartsWith("LEVEL")) // LEVEL ile başlayan sütunları bul
                .Select(c => new
                {
                    Level = int.TryParse(c.Name.Replace("LEVEL ", ""), out int level) ? level : -1, // LEVEL x’ten x değerini al
                    ColumnName = c.Name
                })
                .Where(x => x.Level >= 0) // Geçerli Level’leri al
                .OrderByDescending(x => x.Level) // En yüksek Level’dan başlayarak sırala
                .ToList();

            // Eğer hiç LEVEL sütunu yoksa işlemi durdur
            if (!mevcutLevelSutunlari.Any()) return;

            // Teslim tarihlerini saklayacağımız sözlük
            var teslimTarihleri = new Dictionary<string, string>();

            // DataGridView'deki verileri al
            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                // Teslim tarihini al
                string teslimTarihi = row.Cells["TESLİM TARİHİ"].Value?.ToString();

                // Tüm mevcut Level sütunlarını kontrol et (En yüksek Level’dan başlayarak)
                for (int i = 0; i < mevcutLevelSutunlari.Count; i++)
                {
                    var currentLevel = mevcutLevelSutunlari[i]; // Şu anki Level sütunu
                    var nextLevel = (i + 1 < mevcutLevelSutunlari.Count) ? mevcutLevelSutunlari[i + 1] : null; // Bir alt Level varsa onu al

                    string hucreKodu = row.Cells[currentLevel.ColumnName].Value?.ToString(); // O hücredeki kod
                    string sonrakiLevel = nextLevel != null ? row.Cells[nextLevel.ColumnName].Value?.ToString() : null; // Bir alt Level hücresi

                    if (!string.IsNullOrEmpty(hucreKodu)) // Hücre doluysa
                    {
                        // Eğer teslim tarihi yoksa, önceki seviyeden al
                        if (!teslimTarihleri.ContainsKey(hucreKodu))
                        {
                            teslimTarihleri[hucreKodu] = teslimTarihi;
                        }
                    }
                }
            }

            // Şimdi en düşük Level'dan (Level 0) başlayarak eksik teslim tarihlerini tamamlayalım
            foreach (var level in mevcutLevelSutunlari.OrderBy(x => x.Level)) // Level 0'dan itibaren sırala
            {
                foreach (DataGridViewRow row in advancedDataGridView1.Rows)
                {
                    string hucreKodu = row.Cells[level.ColumnName].Value?.ToString(); // O Level'daki kod

                    if (!string.IsNullOrEmpty(hucreKodu))
                    {
                        // Eğer bu kod için bir teslim tarihi atanmışsa, alt Level'lara da ekle
                        if (teslimTarihleri.ContainsKey(hucreKodu))
                        {
                            string teslimTarihi = teslimTarihleri[hucreKodu];

                            // Alt Level'daki tüm kodları kontrol et ve eksik tarihleri tamamla
                            foreach (var altLevel in mevcutLevelSutunlari.Where(l => l.Level > level.Level))
                            {
                                string altKod = row.Cells[altLevel.ColumnName].Value?.ToString();
                                if (!string.IsNullOrEmpty(altKod) && !teslimTarihleri.ContainsKey(altKod))
                                {
                                    teslimTarihleri[altKod] = teslimTarihi;
                                }
                            }
                        }
                    }
                }
            }

            // Eşleştir listesi ile teslim tarihlerini birleştir
            var yeniListe = eşleştir
                .Select(e => new
                {
                    e.item_partnumber,
                    e.item_sirano,
                    e.title,
                    e.fason,
                    e.rotaitemkod1,
                    e.material,
                    e.hammadde_erp_kodu,
                    TeslimTarihi = teslimTarihleri.ContainsKey(e.item_partnumber) &&
                       DateTime.TryParse(teslimTarihleri[e.item_partnumber]?.ToString(), out DateTime parsedDate)
                       ? (parsedDate < DateTime.Now.Date ? DateTime.Now.Date : parsedDate)
                       : DateTime.Now.Date
                })
                .ToList();


            // transformedList'i Substring(0,13) + Substring(14,2) sırasına göre düzenleyelim
            var sortedList = transformedList
                .OrderBy(t => t.item_partnumber.Substring(0, 13))  // Ana rotaya göre sırala
                .ThenBy(t => int.TryParse(t.item_partnumber.Substring(14, 2), out int num) ? num : int.MaxValue) // 01, 02, 03'e göre sırala
                .ToList();

            // transformedList ile yeniListe'yi birleştir ve teslim tarihini ekle
            var combinedList = sortedList.Select(t =>
            {
                var match = (t.item_partnumber.StartsWith("02.") || t.item_partnumber.StartsWith("01.") || t.item_partnumber.StartsWith("99."))
                    ? yeniListe.FirstOrDefault(y => y.item_partnumber.StartsWith(t.item_partnumber.Substring(0, 13)))
                    : yeniListe.FirstOrDefault(y => y.item_partnumber == t.item_partnumber);

                // Teslim tarihini DateTime olarak al
                DateTime? rotaTeslimTarihi = null;
                if (match != null && DateTime.TryParse(match.TeslimTarihi.ToString(), out DateTime parsedDate))
                {
                    rotaTeslimTarihi = parsedDate;
                }

                // sto_siparis_sure alanını int olarak al
                int? siparisSure = null;
                if (int.TryParse(t.sto_siparis_sure?.ToString(), out int parsedSure))
                {
                    siparisSure = parsedSure;
                }

                return new
                {
                    item_partnumber = t.item_partnumber,
                    sto_siparis_sure = siparisSure,
                    teslimTarihi = rotaTeslimTarihi // İlk teslim tarihi burada belirleniyor
                };
            }).ToList();

            // Teslim tarihlerini güncelle
            var updatedList = new List<(string item_partnumber, int? sto_siparis_sure, DateTime? teslimTarihi)>();

            for (int i = 0; i < combinedList.Count; i++)
            {
                var currentItem = combinedList[i];
                DateTime? yeniTeslimTarihi = currentItem.teslimTarihi; // Başlangıçta mevcut teslim tarihini al

                // Ana rotayı bul (örneğin: 02.03.22.0001)
                string mainRota = currentItem.item_partnumber.Substring(0, 13); // İlk 13 karakter ana rota

                // Önceki öğenin teslim tarihini bul
                if (i > 0) // Eğer ilk öğe değilse
                {
                    var previousItem = updatedList[i - 1]; // Önceki öğe
                    if (previousItem.item_partnumber.StartsWith(mainRota))
                    {
                        // Önceki öğenin teslim tarihini al ve mevcut öğenin sto_siparis_sure'sini ekle
                        if (previousItem.teslimTarihi.HasValue && currentItem.sto_siparis_sure.HasValue)
                        {
                            yeniTeslimTarihi = previousItem.teslimTarihi.Value.AddDays(currentItem.sto_siparis_sure.Value);
                        }
                    }
                }

                // Güncellenmiş nesneyi listeye ekle
                updatedList.Add((currentItem.item_partnumber, currentItem.sto_siparis_sure, yeniTeslimTarihi));
            }

            // PublicVeriListesi'ni nesneye çevir
            var publicVeriListesiObjeleri = PublicVeriListesi
                .Select(item =>
                {
                    var parts = item.Split(" - "); // "-" ile ayır
                    return new
                    {
                        ProjeKod = parts[0], // İlk parça: Proje Kodu
                        IsEmriKodu = parts[1], // İkinci parça: İş Emri Kodu
                        UretilecekUrunKodu = parts[2] // Üçüncü parça: Üretilecek Ürün Kodu
                    };
                }).ToList();

            // Sonuç formatını ayarla
            var finalList = updatedList
                .Select(item =>
                {
                    // Varsayılan iş emri kodunu null olarak başlat
                    string isEmriKodu = null;

                    // item_partnumber ile UretilecekUrunKodu eşleşiyorsa iş emri kodunu al
                    var matchingItem = publicVeriListesiObjeleri
                        .FirstOrDefault(publicItem => publicItem.UretilecekUrunKodu == item.item_partnumber);

                    if (matchingItem != null)
                    {
                        isEmriKodu = matchingItem.IsEmriKodu; // Eşleşen iş emri kodunu al
                    }

                    // Eşleşen iş emri kodu varsa yeni nesneyi döndür, yoksa null döndür
                    return new
                    {
                        item.item_partnumber,
                        item.sto_siparis_sure,
                        teslimTarihi = item.teslimTarihi.HasValue ? item.teslimTarihi.Value.ToString("dd.MM.yyyy") : null,
                        IsEmriKodu = isEmriKodu // İş emri kodunu ekle
                    };
                })
                .Where(result => result.IsEmriKodu != null) // IsEmriKodu null olmayanları filtrele
                .ToList();

            // finalList içindeki öğeler üzerinden URETIM_ROTA_PLANLARI sorgusu
            foreach (var item in finalList)
            {
                // URETIM_ROTA_PLANLARI tablosundan verileri al
                var matchingPlans = dbContext.URETIM_ROTA_PLANLARI
                    .Where(plan => plan.RtP_IsEmriKodu == item.IsEmriKodu && // IsEmriKodu eşleşmesi
                                   plan.RtP_UrunKodu == item.item_partnumber) // item_partnumber eşleşmesi
                    .ToList(); // Sonuçları listeye al

                // Eşleşen her bir kaydı güncelle
                foreach (var plan in matchingPlans)
                {
                    // Planlanan Başlama Tarihi'ni güncelle
                    // Eğer item.teslimTarihi boş değilse ve tarih formatı uygunsa
                    if (DateTime.TryParse(item.teslimTarihi, out DateTime parsedDate))
                    {
                        plan.Rtp_PlanlananBaslamaTarihi = parsedDate; // DateTime olarak atama
                    }
                    else
                    {
                        // Hata yönetimi: Geçersiz tarih durumu
                        // Gerekirse bir hata mesajı yazdırabilirsin
                        Console.WriteLine($"Geçersiz tarih: {item.teslimTarihi}");
                    }
                }
            }
            // finalList içindeki öğeler üzerinden URETIM_ROTA_PLANLARI sorgusu
            foreach (var item in finalList)
            {
                // URETIM_ROTA_PLANLARI tablosundan verileri al
                var matchingPlans = dbContext.ISEMIRLERI
                    .Where(plan => plan.is_Kod == item.IsEmriKodu) // item_partnumber eşleşmesi
                    .ToList(); // Sonuçları listeye al

                // Eşleşen her bir kaydı güncelle
                foreach (var plan in matchingPlans)
                {
                    // Planlanan Başlama Tarihi'ni güncelle
                    // Eğer item.teslimTarihi boş değilse ve tarih formatı uygunsa
                    if (DateTime.TryParse(item.teslimTarihi, out DateTime parsedDate))
                    {
                        plan.is_Emri_PlanBaslamaTarihi = parsedDate; // DateTime olarak atama
                    }
                    else
                    {
                        // Hata yönetimi: Geçersiz tarih durumu
                        // Gerekirse bir hata mesajı yazdırabilirsin
                        Console.WriteLine($"Geçersiz tarih: {item.teslimTarihi}");
                    }
                }
            }
            // Değişiklikleri kaydet
            dbContext.SaveChanges();
            MessageBox.Show("Başarıyla Güncellendi...");

        }
    }
}
