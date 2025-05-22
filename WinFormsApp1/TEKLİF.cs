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
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using WinFormsApp1.Tables;
using SİPARİŞ_GİRİŞ.Tables;
using ExcelDataReader;
using System.IO;

namespace SİPARİŞ_GİRİŞ
{
    public partial class TEKLİF : Form
    {
        private MyDbContext dbContext;
        public TEKLİF()
        {
            dbContext = new MyDbContext();
            InitializeComponent();
        }
        //01-12-01-02=dilek taçyıldız
        List<Dictionary<string, object>> customerList = new List<Dictionary<string, object>>();
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog file = new OpenFileDialog();
                file.Filter = "Excel Dosyası |*.xlsx;*.xls";
                if (file.ShowDialog() == DialogResult.OK)
                {
                    string tamYol = file.FileName;

                    // Excel dosyasını okuyun
                    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                    using (var stream = File.Open(tamYol, FileMode.Open, FileAccess.Read))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                                {
                                    UseHeaderRow = true
                                }
                            });

                            if (result.Tables.Count > 0)
                            {
                                DataTable data = result.Tables[0];

                                // customerList'i temizleyin
                                customerList.Clear();

                                // İstediğin 4 sütunla yeni DataTable oluştur
                                DataTable filteredData = new DataTable();
                                filteredData.Columns.Add("Part Number");
                                filteredData.Columns.Add("Item QTY");
                                filteredData.Columns.Add("Title");
                                filteredData.Columns.Add("Item");

                                foreach (DataRow row in data.Rows)
                                {
                                    string itemValue = row["Item"]?.ToString();

                                    if (!string.IsNullOrEmpty(itemValue) && !itemValue.Contains("."))
                                    {
                                        // Yeni satır oluştur ve sadece seçilen sütunları al
                                        DataRow newRow = filteredData.NewRow();
                                        newRow["Part Number"] = row["Part Number"];
                                        newRow["Item QTY"] = row["Item QTY"];
                                        newRow["Title"] = row["Title"];
                                        newRow["Item"] = row["Item"];
                                        filteredData.Rows.Add(newRow);

                                        // customerList'e de bu bilgileri ekle
                                        Dictionary<string, object> item = new Dictionary<string, object>();
                                        item["Part Number"] = row["Part Number"];
                                        item["Item QTY"] = row["Item QTY"];
                                        item["Title"] = row["Title"];
                                        item["Item"] = row["Item"];
                                        customerList.Add(item);
                                    }
                                }

                                // AdvancedDataGridView'e sadece bu sütunları yükle
                                advancedDataGridView1.DataSource = filteredData;

                                MessageBox.Show("BOM Listesi Yüklendi (sadece seçili sütunlar, noktalar hariç)");
                            }
                            else
                            {
                                MessageBox.Show("Excel dosyasında veri bulunamadı.");
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Dosya seçilmedi.");
                }
            }
            catch (IOException ioEx)
            {
                MessageBox.Show("Dosya hatası: " + ioEx.Message);
            }
            catch (FormatException formatEx)
            {
                MessageBox.Show("Format hatası: " + formatEx.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Bir hata oluştu: " + ex.Message);
            }



        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<VERILEN_TEKLIFLER> verilenteklif = new List<VERILEN_TEKLIFLER>();
            int nextStlEvrakSira = dbContext.VERILEN_TEKLIFLER.Max(item => item.tkl_evrakno_sira);
            int evrakSira = nextStlEvrakSira + 1; // Benzersiz bir şekilde ayarlanmalıdır

            string projekodu = textBox2.Text;
            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                if (!row.IsNewRow)
                {
                    // DataGridView'den verileri al
                    string parçano = row.Cells["Part Number"].Value.ToString();
                    float Adet = Convert.ToSingle(row.Cells["Item QTY"].Value);
                    string parçaismi = row.Cells["Title"].Value.ToString();
                    int satırno = Convert.ToInt32(row.Cells["Item"].Value);

                    // Benzersiz değerleri ayarlayın
                    string evrakSeri = "VT"; // Örnek olarak "T" olarak ayarlandı, siz iş mantığınıza göre değiştirmelisiniz
                    VERILEN_TEKLIFLER talep1 = new VERILEN_TEKLIFLER
                    {
                        tkl_Guid = Guid.NewGuid(), // Benzersiz bir GUID oluştur
                        tkl_DBCno = 0,
                        tkl_SpecRECno = 0,
                        tkl_iptal = 0,
                        tkl_fileid = 100,
                        tkl_hidden = 0,
                        tkl_kilitli = 0,
                        tkl_degisti = 0,
                        tkl_checksum = 0,
                        tkl_create_user = 32,
                        tkl_create_date = DateTime.Now,
                        tkl_lastup_user = 32,
                        tkl_lastup_date = DateTime.Now,
                        tkl_special1 = "",
                        tkl_special2 = "",
                        tkl_special3 = "",
                        tkl_firmano = 0,
                        tkl_subeno = 0,
                        tkl_stok_kod = parçano,
                        tkl_cari_kod = "",
                        tkl_evrakno_seri = evrakSeri,
                        tkl_evrakno_sira = evrakSira,
                        tkl_evrak_tarihi= DateTime.Now.Date,
                        tkl_satirno = satırno,
                        tkl_belge_no = "",
                        tkl_belge_tarih = DateTime.Now.Date,
                        tkl_asgari_miktar = 0,
                        tkl_teslimat_suresi = 0,
                        tkl_baslangic_tarihi = DateTime.Now.Date,
                        tkl_Gecerlilik_Sures = DateTime.Now.Date,
                        tkl_Brut_fiyat = 0,
                        tkl_Odeme_Plani = 0,
                        tkl_Alisfiyati = 0,
                        tkl_karorani = 0,
                        tkl_miktar = Adet,
                        tkl_Aciklama = "",
                        tkl_doviz_cins = 0,
                        tkl_doviz_kur = 1,
                        tkl_alt_doviz_kur = 0,
                        tkl_iskonto1 = 0,
                        tkl_iskonto2 = 0,
                        tkl_iskonto3 = 0,
                        tkl_iskonto4 = 0,
                        tkl_iskonto5 = 0,
                        tkl_iskonto6 = 0,
                        tkl_masraf1 = 0,
                        tkl_masraf2 = 0,
                        tkl_masraf3 = 0,
                        tkl_masraf4 = 0,
                        tkl_vergi_pntr = 0,
                        tkl_vergi = 0,
                        tkl_masraf_vergi_pnt = 0,
                        tkl_masraf_vergi = 0,
                        tkl_isk_mas1 = 0,
                        TKL_ISK_MAS2 = 1,
                        TKL_ISK_MAS3 = 1,
                        TKL_ISK_MAS4 = 1,
                        TKL_ISK_MAS5 = 1,
                        TKL_ISK_MAS6 = 1,
                        TKL_ISK_MAS7 = 1,
                        TKL_ISK_MAS8 = 1,
                        TKL_ISK_MAS9 = 1,
                        TKL_ISK_MAS10 = 1,
                        TKL_SAT_ISKMAS1 = 0,
                        TKL_SAT_ISKMAS2 = 0,
                        TKL_SAT_ISKMAS3 = 0,
                        TKL_SAT_ISKMAS4 = 0,
                        TKL_SAT_ISKMAS5 = 0,
                        TKL_SAT_ISKMAS6 = 0,
                        TKL_SAT_ISKMAS7 = 0,
                        TKL_SAT_ISKMAS8 = 0,
                        TKL_SAT_ISKMAS9 = 0,
                        TKL_SAT_ISKMAS10 = 0,
                        TKL_VERGISIZ_FL = 0,
                        TKL_KAPAT_FL = 0,
                        TKL_TESLIMTURU = "",
                        tkl_ProjeKodu = projekodu,
                        tkl_Sorumlu_Kod = "01-12-01-02",
                        tkl_adres_no = 0,
                        tkl_yetkili_uid = Guid.Empty, // Varsayılan GUID değeri
                        tkl_durumu = 0,
                        tkl_TedarikEdilecekCari = "",
                        tkl_fiyat_liste_no = 0,
                        tkl_Birimfiyati = 0,
                        tkl_paket_kod = "",
                        tkl_teslim_miktar = 0,
                        tkl_OnaylayanKulNo = 0,
                        tkl_cagrilabilir_fl = true,
                        tkl_harekettipi = 0,
                        tkl_cari_sormerk = projekodu,
                        tkl_stok_sormerk = projekodu,
                        tkl_kapatmanedenkod = "",
                        tkl_servisisemrikodu = "",
                        tkl_birim_pntr = 1,
                        tkl_cari_tipi = 0,
                        tkl_HareketGrupKodu1 = "",
                        tkl_HareketGrupKodu2 = "",
                        tkl_HareketGrupKodu3 = "",
                        tkl_Olcu1 = 0,
                        tkl_Olcu2 = 0,
                        tkl_Olcu3 = 0,
                        tkl_Olcu4 = 0,
                        tkl_Olcu5 = 0,
                        tkl_FormulMiktarNo = 0,
                        tkl_FormulMiktar = 0,
                        tkl_Tevkifat_turu = 0,
                        tkl_tevkifat_sifirlandi_fl = false
                    };

                    verilenteklif.Add(talep1);
                }
            }

            // Veritabanına ekle
            dbContext.VERILEN_TEKLIFLER.AddRange(verilenteklif);
            dbContext.SaveChanges(); // Değişiklikleri kaydet

            MessageBox.Show("KAYDEDİLDİ...");

        }
    }
}
