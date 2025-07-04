using Castle.Core.Resource;
using SÝPARÝÞ_GÝRÝÞ;
using SÝPARÝÞ_GÝRÝÞ.Tables;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Security.Cryptography;
using WinFormsApp1.Context;
using WinFormsApp1.Tables;

namespace WinFormsApp1
{
    public partial class Form1 : Form
    {
        private MyDbContext dbContext;
        public Form1()
        {
            dbContext = new MyDbContext();

            InitializeComponent();

            DataGridViewCheckBoxColumn checkBoxColumn1 = new DataGridViewCheckBoxColumn();
            checkBoxColumn1.HeaderText = "T";
            checkBoxColumn1.Name = "checkBoxColumn1";
            advancedDataGridView1.Columns.Insert(0, checkBoxColumn1);
        }
        DataTable table = new DataTable();
        // Verileri bir listeye aktar
        List<Dictionary<string, object>> customerList = new List<Dictionary<string, object>>();
        public void siparisGiris()
        {
            table.Columns.Clear();
            table.Rows.Clear();
            table.Clear();
            var selectedDate1 = dateTimePicker2.Value.Date;
            var selectedDate = dateTimePicker1.Value.Date;
            int Durum = Convert.ToInt32(comboBox1.Text.ToString());
            // Define columns in the DataGridView
            table.Columns.Add("SATIR_NO", typeof(int));
            table.Columns.Add("PROJE_KOD", typeof(string));
            table.Columns.Add("ÝÞ_EMRÝ_KODU", typeof(string));
            table.Columns.Add("SÝPARÝÞ_SERÝ", typeof(string));
            table.Columns.Add("SÝPARÝÞ_NO", typeof(string));
            table.Columns.Add("ÜRETÝLECEK_ÜRÜN_KODU", typeof(string));
            table.Columns.Add("ÜRETÝLECEK_ÜRÜN_ADI", typeof(string));
            table.Columns.Add("ÜRETÝLECEK_ÜRÜN_MÝKTARI", typeof(double));
            table.Columns.Add("TÜKETÝLECEK_ÜRÜN_KODU_AÇIKLAMA", typeof(string));
            table.Columns.Add("TÜKETÝLECEK_ÜRÜN_KODU", typeof(string));
            table.Columns.Add("TÜKETÝLECEK_ÜRÜN_ADI", typeof(string));
            table.Columns.Add("TÜKETÝLECEK_ÜRÜN_MÝKTARI", typeof(double));
            table.Columns.Add("ÝSEMRÝ_TÝPÝ", typeof(string));
            table.Columns.Add("DURUMU", typeof(byte));
            table.Columns.Add("TARÝH", typeof(DateTime));


            var sorgu = from isEmir in dbContext.ISEMIRLERI
                        join uretim in dbContext.URETIM_MALZEME_PLANLAMA on isEmir.is_Kod equals uretim.upl_isemri
                        where isEmir.is_Kod.Substring(0, 8) == textBox1.Text &&
                              isEmir.is_create_date.Date >= selectedDate &&
                              isEmir.is_create_date.Date <= selectedDate1 &&
                              isEmir.is_EmriDurumu == Durum && // Durum kontrolü
                              dbContext.ISEMIRLERI_USER.Any(u => u.Record_uid == isEmir.is_Guid && u.is_emri_tipi != null) && // Null kontrolü
                              uretim.upl_satirno != 0
                        orderby uretim.upl_kodu ascending
                        select new
                        {
                            SATIR_NO = uretim.upl_satirno,
                            PROJE_KOD = isEmir.is_ProjeKodu,
                            ÝÞ_EMRÝ_KODU = isEmir.is_Kod,
                            SÝPARÝÞ_SERÝ = isEmir.is_SiparisNo_Seri,
                            SÝPARÝÞ_NO = isEmir.is_SiparisNo_Sira,
                            ÜRETÝLECEK_ÜRÜN_KODU = uretim.upl_urstokkod,
                            ÜRETÝLECEK_ÜRÜN_ADI = dbContext.STOKLAR.FirstOrDefault(u => u.sto_kod == uretim.upl_urstokkod).sto_isim,
                            ÜRETÝLECEK_ÜRÜN_MÝKTARI = uretim.upl_uret_miktar,
                            TÜKETÝLECEK_ÜRÜN_KODU_AÇIKLAMA = uretim.upl_aciklama,
                            TÜKETÝLECEK_ÜRÜN_KODU = uretim.upl_kodu,
                            TÜKETÝLECEK_ÜRÜN_ADI = dbContext.STOKLAR.FirstOrDefault(u => u.sto_kod == uretim.upl_kodu).sto_isim,
                            TÜKETÝLECEK_ÜRÜN_MÝKTARI = uretim.upl_sarfmiktari,
                            ÝSEMRÝ_TÝPÝ = dbContext.ISEMIRLERI_USER.FirstOrDefault(u => u.Record_uid == isEmir.is_Guid).is_emri_tipi,
                            DURUMU = isEmir.is_EmriDurumu,
                            TARÝH = isEmir.is_create_date.Date
                        };




            // Sorgu sonucundaki verileri liste olarak alýn
            var veriListesi = sorgu.ToList();



            // Eðer veri yoksa kullanýcýya mesaj gösterin ve DataGridView'i temizleyin
            if (veriListesi.Count == 0)
            {
                MessageBox.Show("Seçilen tarihte veri bulunmamaktadýr.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                advancedDataGridView1.DataSource = null; // DataGridView'i temizle
            }
            else
            {
                // DataGridView baþlýk satýrýnýn görünümünü ayarlayýn
                advancedDataGridView1.EnableHeadersVisualStyles = false;
                advancedDataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.DimGray;
                advancedDataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

                // Özel bir yazý tipi tanýmlayýn
                Font baslikYaziTipi = new Font("Arial", 10, FontStyle.Regular);

                // DataGridView baþlýk satýrýnýn yazý tipini ayarlayýn
                advancedDataGridView1.ColumnHeadersDefaultCellStyle.Font = baslikYaziTipi;
                foreach (var urun in veriListesi)
                {
                    table.Rows.Add(urun.SATIR_NO, urun.PROJE_KOD, urun.ÝÞ_EMRÝ_KODU, urun.SÝPARÝÞ_SERÝ, urun.SÝPARÝÞ_NO, urun.ÜRETÝLECEK_ÜRÜN_KODU, urun.ÜRETÝLECEK_ÜRÜN_ADI, urun.ÜRETÝLECEK_ÜRÜN_MÝKTARI,
                        urun.TÜKETÝLECEK_ÜRÜN_KODU_AÇIKLAMA, urun.TÜKETÝLECEK_ÜRÜN_KODU, urun.TÜKETÝLECEK_ÜRÜN_ADI, urun.TÜKETÝLECEK_ÜRÜN_MÝKTARI, urun.ÝSEMRÝ_TÝPÝ, urun.DURUMU, urun.TARÝH);
                }

                // DataGridView'e verileri yükleyin
                advancedDataGridView1.DataSource = table;
                // DataGridView'e verileri yükledikten sonra CheckBox'larý iþaretle
                for (int rowIndex = 0; rowIndex < advancedDataGridView1.Rows.Count; rowIndex++)
                {
                    var veri = veriListesi[rowIndex]; // Mevcut satýrýn verisi
                    bool shouldCheck = false;

                    // Ýlk koþullarýn kontrolü
                    if ((veri.ÝSEMRÝ_TÝPÝ == "SATINALMA_IE" &&
                         !new[] { ".128" }.Any(prefix => veri.ÜRETÝLECEK_ÜRÜN_KODU.EndsWith(prefix))) ||
                        (veri.ÝSEMRÝ_TÝPÝ == "KK_IE" &&
                         new[] { ".128" }.Any(prefix => veri.TÜKETÝLECEK_ÜRÜN_KODU.EndsWith(prefix))) ||
                        (veri.ÝSEMRÝ_TÝPÝ == "MONTAJ_IE" &&
                         (!dbContext.URUNLER.Any(x => x.uru_stok_kod == veri.TÜKETÝLECEK_ÜRÜN_KODU)))
                    )
                    {
                        shouldCheck = true;
                    }

                    if (veri.ÝSEMRÝ_TÝPÝ == "MONTAJ_IE")
                    {
                        foreach (var x in customerList)
                        {
                            if (
                                x.ContainsKey("Part Number") &&
                                new[] { "02" }.Any(prefix => veri.TÜKETÝLECEK_ÜRÜN_KODU.StartsWith(prefix)) &&
                                x["Part Number"].ToString() == veri.TÜKETÝLECEK_ÜRÜN_KODU &&
                                x.ContainsKey("Fason") &&
                                x["Fason"].ToString() != "EVET" &&
                                x.ContainsKey("Malzeme Taným") &&
                                !string.IsNullOrEmpty(x["Malzeme Taným"].ToString()) &&
                                !new[]
                                {
                                    "T.KESIM - 000",
                                    "LAZERLE KESIM - 120",
                                    "OKSIJEN KESIM - 121",
                                    "ROUTER KESIM - 127",
                                    "SU JETI KESIM - 122"
                                }.Contains(x["Rota01"].ToString())
                            )
                            {
                                shouldCheck = true;
                                break;
                            }
                        }
                    }


                    // Burada shouldCheck deðiþkenini kullanarak gerekli iþlemleri yapabilirsiniz


                    // Eðer koþullar saðlanýyorsa checkbox'ý iþaretle
                    advancedDataGridView1.Rows[rowIndex].Cells["checkBoxColumn1"].Value = shouldCheck;
                }


            }
        }
        private void InsertDataToDatabase()
        {
            var selectedDateEvrakHammade = dateTimePicker3.Value.Date;
            var selectedDateEvrakResSA = dateTimePicker5.Value.Date;
            var selectedDateEvrakComp = dateTimePicker6.Value.Date;
            var selectedDateTesHammade = dateTimePicker10.Value.Date;
            var selectedDateTesResSA = dateTimePicker8.Value.Date;
            var selectedDateTesComp = dateTimePicker7.Value.Date;


            Form sebepSecimiForm = new Form
            {
                Text = "Ara Verme Sebebinizi Seçiniz...",
                Size = new Size(300, 200),
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterScreen
            };

            ComboBox comboBox = new ComboBox
            {
                Location = new Point(20, 20),
                Width = 200,
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            comboBox.Items.AddRange(new string[]
            {
                "29",
                "15",
                "31",
                "25",
                "28",
                "34"
            });

            TextBox textBox = new TextBox
            {
                Location = new Point(20, 60),
                Width = 200,
                Visible = false  // Baþlangýçta görünmez
            };
            TextBox textBoxBelge = new TextBox
            {
                Location = new Point(50, 100),
                Width = 200,
                Visible = false  // Baþlangýçta görünmez
            };
            Button onayButton = new Button
            {
                Text = "Onay",
                Location = new Point(100, 100),
                Size = new Size(60, 30),
            };
            // ComboBox'ýn SelectedIndexChanged etkinliði
            comboBox.SelectedIndexChanged += (sender, e) =>
            {
                int selectedIndex = comboBox.SelectedIndex;
                if (selectedIndex == 0)
                {
                    textBox.Text = "01-09-02-11";
                    textBoxBelge.Text = "SH";
                }
                else if (selectedIndex == 1)
                {
                    textBox.Text = "01-08-01-01";
                    textBoxBelge.Text = "SÞ";
                }
                else if (selectedIndex == 2)
                {
                    textBox.Text = "01-09-02-03";
                    textBoxBelge.Text = "MK";
                }
                else if (selectedIndex == 3)
                {
                    textBox.Text = "01-08-02-06";
                    textBoxBelge.Text = "TK";
                }
                else if (selectedIndex == 4)
                {
                    textBox.Text = "01-08-02-04";
                    textBoxBelge.Text = "ÞB";
                }
                else if (selectedIndex == 5)
                {
                    textBox.Text = "01-09-02-15";
                    textBoxBelge.Text = "SB";
                }
            };
            onayButton.Click += (s, e) =>
        {
            if (textBox1.Text.StartsWith("3S") || textBox1.Text.StartsWith("2S") || textBox1.Text.StartsWith("XS") || textBox1.Text.StartsWith("OX") || textBox1.Text.StartsWith("24"))
            {
                List<SATINALMA_TALEPLERI> talepListesiHamMadde = new List<SATINALMA_TALEPLERI>();
                List<SATINALMA_TALEPLERI> talepListesiFason = new List<SATINALMA_TALEPLERI>();
                List<SATINALMA_TALEPLERI> talepListesiComponent = new List<SATINALMA_TALEPLERI>();
                List<SATINALMA_TALEPLERI> talepListesiResimliSA = new List<SATINALMA_TALEPLERI>();
                int nextStlEvrakSira = dbContext.SATINALMA_TALEPLERI.Max(item => item.stl_evrak_sira);
                int evrakSira = nextStlEvrakSira + 1; // Benzersiz bir þekilde ayarlanmalýdýr
                                                      // DataGridView'deki her bir satýr için
                int sayac = 0;
                int satirNo = sayac;
                foreach (DataGridViewRow row in advancedDataGridView1.Rows)
                {


                    if (!row.IsNewRow)
                    {

                        // DataGridView'den verileri al
                        string projeKodu = row.Cells["PROJE_KOD"].Value.ToString();
                        string isEmriKodu = row.Cells["ÝÞ_EMRÝ_KODU"].Value.ToString();
                        string üretilecekUrunKodu = row.Cells["ÜRETÝLECEK_ÜRÜN_KODU"].Value.ToString();
                        double üretilecekUrunMiktari = double.Parse(row.Cells["ÜRETÝLECEK_ÜRÜN_MÝKTARI"].Value.ToString());
                        string tüketilecekUrunKodu = row.Cells["TÜKETÝLECEK_ÜRÜN_KODU"].Value.ToString();
                        double tüketilecekUrunMiktari = double.Parse(row.Cells["TÜKETÝLECEK_ÜRÜN_MÝKTARI"].Value.ToString());
                        string tüketilecekUrunKoduAciklama = row.Cells["TÜKETÝLECEK_ÜRÜN_KODU_AÇIKLAMA"].Value.ToString();
                        string isemritipi = row.Cells["ÝSEMRÝ_TÝPÝ"].Value.ToString();
                        string siparisseri = row.Cells["SÝPARÝÞ_SERÝ"].Value.ToString();
                        string siparissýra = row.Cells["SÝPARÝÞ_NO"].Value.ToString();
                        // Benzersiz deðerleri ayarlayýn
                        string evrakSeri = "T"; // Örnek olarak "T" olarak ayarlandý, siz iþ mantýðýnýza göre deðiþtirmelisiniz

                        string SipSeriSýra = siparisseri + "-" + siparissýra;
                        string[] allowedEndings = new string[] { ".120", ".121", ".122", ".127" };
                        string[] allowedEndings1 = new string[] { ".000" };
                        string[] allowedEndings2 = new string[] { ".128" };
                        Dictionary<string, string> numaraKarsilik = new Dictionary<string, string>
                        {
                                { ".120", "LAZER KESÝM" },
                                { ".121", "OKSÝJEN KESÝM" },
                                { ".122", "SU JETÝ KESÝM" },
                                { ".127", "ROUTER KESÝM" }
                        };
                        var rsa = dbContext.URETIM_MALZEME_PLANLAMA.Where(x => x.upl_kodu == üretilecekUrunKodu).Select(x => x.upl_urstokkod).FirstOrDefault();
                        //HAMMADDE-KESÝMLER
                        if (Convert.ToBoolean(row.Cells["checkBoxColumn1"].Value) && allowedEndings.Any(end => üretilecekUrunKodu.EndsWith(end)))
                        {
                            string numara = allowedEndings.FirstOrDefault(end => üretilecekUrunKodu.EndsWith(end));
                            if (numara != null && numaraKarsilik.ContainsKey(numara))
                            {
                                string aciklama = üretilecekUrunMiktari + " Adet /" + numaraKarsilik[numara];

                                SATINALMA_TALEPLERI talep1 = new SATINALMA_TALEPLERI
                                {
                                    stl_DBCno = 0,
                                    stl_SpecRECno = 0,
                                    stl_iptal = 0,
                                    stl_fileid = 118,
                                    stl_hidden = 0,
                                    stl_kilitli = 0,
                                    stl_degisti = 0,
                                    stl_checksum = 0,
                                    stl_create_user = comboBox.Text,
                                    stl_create_date = DateTime.Now,
                                    stl_lastup_user = comboBox.Text,
                                    stl_lastup_date = DateTime.Now,
                                    stl_special1 = comboBox.Text,
                                    stl_special2 = "",
                                    stl_special3 = "",
                                    stl_firmano = 0,
                                    stl_subeno = 0,
                                    stl_tarihi = selectedDateEvrakHammade,
                                    stl_teslim_tarihi = selectedDateTesHammade,
                                    stl_evrak_seri = evrakSeri,
                                    stl_evrak_sira = evrakSira,
                                    stl_satir_no = satirNo,
                                    stl_belge_no = textBoxBelge.Text,
                                    stl_belge_tarihi = DateTime.Now.Date,
                                    stl_Sor_Merk = SipSeriSýra,
                                    stl_Stok_kodu = tüketilecekUrunKodu,
                                    stl_Satici_Kodu = "",
                                    stl_miktari = tüketilecekUrunMiktari,
                                    stl_teslim_miktari = 0,
                                    stl_aciklama = aciklama + "&" + textBox1.Text,
                                    stl_aciklama2 = üretilecekUrunKodu,
                                    stl_depo_no = 70150,
                                    stl_kapat_fl = 0,
                                    stl_projekodu = projeKodu,
                                    stl_parti_kodu = "",
                                    stl_lot_no = "0",
                                    stl_OnaylayanKulNo = "0",
                                    stl_cagrilabilir_fl = 1,
                                    stl_harekettipi = 0,
                                    stl_talep_eden = textBox.Text,
                                    stl_kapatmanedenkod = "",
                                    stl_KaynakSip_uid = Guid.Empty,
                                    stl_HareketGrupKodu1 = isEmriKodu,
                                    stl_HareketGrupKodu2 = "",
                                    stl_HareketGrupKodu3 = "",
                                    stl_birim_pntr = 1,
                                    stl_Franchise_Fiyati = 0,
                                };
                                sayac++;
                                satirNo = sayac;
                                talepListesiHamMadde.Add(talep1);
                            }
                        }
                        //HAMMADDE-DEMÝRÇELÝK
                        else if (Convert.ToBoolean(row.Cells["checkBoxColumn1"].Value) && allowedEndings1.Any(end => üretilecekUrunKodu.EndsWith(end)))
                        {
                            string aciklama = tüketilecekUrunKoduAciklama + " / " + üretilecekUrunMiktari + " Adet";

                            SATINALMA_TALEPLERI talep1 = new SATINALMA_TALEPLERI
                            {
                                stl_DBCno = 0,
                                stl_SpecRECno = 0,
                                stl_iptal = 0,
                                stl_fileid = 118,
                                stl_hidden = 0,
                                stl_kilitli = 0,
                                stl_degisti = 0,
                                stl_checksum = 0,
                                stl_create_user = comboBox.Text,
                                stl_create_date = DateTime.Now,
                                stl_lastup_user = comboBox.Text,
                                stl_lastup_date = DateTime.Now,
                                stl_special1 = comboBox.Text,
                                stl_special2 = "",
                                stl_special3 = "",
                                stl_firmano = 0,
                                stl_subeno = 0,
                                stl_tarihi = selectedDateEvrakHammade,
                                stl_teslim_tarihi = selectedDateTesHammade,
                                stl_evrak_seri = evrakSeri,
                                stl_evrak_sira = evrakSira,
                                stl_satir_no = satirNo,
                                stl_belge_no = textBoxBelge.Text,
                                stl_belge_tarihi = DateTime.Now.Date,
                                stl_Sor_Merk = projeKodu,
                                stl_Stok_kodu = tüketilecekUrunKodu,
                                stl_Satici_Kodu = "",
                                stl_miktari = tüketilecekUrunMiktari,
                                stl_teslim_miktari = 0,
                                stl_aciklama = aciklama + "&" + textBox1.Text,
                                stl_aciklama2 = üretilecekUrunKodu,
                                stl_depo_no = 70150,
                                stl_kapat_fl = 0,
                                stl_projekodu = projeKodu,
                                stl_parti_kodu = "",
                                stl_lot_no = "0",
                                stl_OnaylayanKulNo = "0",
                                stl_cagrilabilir_fl = 1,
                                stl_harekettipi = 0,
                                stl_talep_eden = textBox.Text,
                                stl_kapatmanedenkod = "",
                                stl_KaynakSip_uid = Guid.Empty,
                                stl_HareketGrupKodu1 = isEmriKodu,
                                stl_HareketGrupKodu2 = "",
                                stl_HareketGrupKodu3 = "",
                                stl_birim_pntr = 1,
                                stl_Franchise_Fiyati = 0,
                            };
                            sayac++;
                            satirNo = sayac;
                            talepListesiHamMadde.Add(talep1);

                        }
                        //RESÝMLÝ SATIN ALMA
                        else if (Convert.ToBoolean(row.Cells["checkBoxColumn1"].Value) && allowedEndings2.Any(end => tüketilecekUrunKodu.EndsWith(end)))
                        {
                            SATINALMA_TALEPLERI talep3 = new SATINALMA_TALEPLERI
                            {
                                stl_DBCno = 0,
                                stl_SpecRECno = 0,
                                stl_iptal = 0,
                                stl_fileid = 118,
                                stl_hidden = 0,
                                stl_kilitli = 0,
                                stl_degisti = 0,
                                stl_checksum = 0,
                                stl_create_user = comboBox.Text,
                                stl_create_date = DateTime.Now,
                                stl_lastup_user = comboBox.Text,
                                stl_lastup_date = DateTime.Now,
                                stl_special1 = comboBox.Text,
                                stl_special2 = "",
                                stl_special3 = "",
                                stl_firmano = 0,
                                stl_subeno = 0,
                                stl_tarihi = selectedDateEvrakResSA,
                                stl_teslim_tarihi = selectedDateTesResSA,
                                stl_evrak_seri = evrakSeri,
                                stl_evrak_sira = evrakSira + 1,
                                stl_satir_no = satirNo,
                                stl_belge_no = textBoxBelge.Text,
                                stl_belge_tarihi = DateTime.Now.Date,
                                stl_Sor_Merk = SipSeriSýra,
                                stl_Stok_kodu = üretilecekUrunKodu,
                                stl_Satici_Kodu = "",
                                stl_miktari = üretilecekUrunMiktari,
                                stl_teslim_miktari = 0,
                                stl_aciklama = rsa,
                                stl_aciklama2 = "",
                                stl_depo_no = 70150,
                                stl_kapat_fl = 0,
                                stl_projekodu = projeKodu,
                                stl_parti_kodu = "",
                                stl_lot_no = "0",
                                stl_OnaylayanKulNo = "0",
                                stl_cagrilabilir_fl = 1,
                                stl_harekettipi = 0,
                                stl_talep_eden = textBox.Text,
                                stl_kapatmanedenkod = "",
                                stl_KaynakSip_uid = Guid.Empty,
                                stl_HareketGrupKodu1 = isEmriKodu,
                                stl_HareketGrupKodu2 = "",
                                stl_HareketGrupKodu3 = "",
                                stl_birim_pntr = 1,
                                stl_Franchise_Fiyati = 0,
                            };
                            sayac++;
                            satirNo = sayac;
                            talepListesiResimliSA.Add(talep3);
                        }
                        // MONTAJ ÝÞ EMRÝ VE COMPANENT ÝÞ EMRÝ OLABÝLÝCEKLER.
                        else if (Convert.ToBoolean(row.Cells["checkBoxColumn1"].Value))
                        {
                            SATINALMA_TALEPLERI talep4 = new SATINALMA_TALEPLERI
                            {
                                stl_DBCno = 0,
                                stl_SpecRECno = 0,
                                stl_iptal = 0,
                                stl_fileid = 118,
                                stl_hidden = 0,
                                stl_kilitli = 0,
                                stl_degisti = 0,
                                stl_checksum = 0,
                                stl_create_user = comboBox.Text,
                                stl_create_date = DateTime.Now,
                                stl_lastup_user = comboBox.Text,
                                stl_lastup_date = DateTime.Now,
                                stl_special1 = comboBox.Text,
                                stl_special2 = "",
                                stl_special3 = "",
                                stl_firmano = 0,
                                stl_subeno = 0,
                                stl_tarihi = selectedDateEvrakComp,
                                stl_teslim_tarihi = selectedDateTesComp,
                                stl_evrak_seri = evrakSeri,
                                stl_evrak_sira = evrakSira + 2,
                                stl_satir_no = satirNo,
                                stl_belge_no = textBoxBelge.Text,
                                stl_belge_tarihi = DateTime.Now.Date,
                                stl_Sor_Merk = SipSeriSýra,
                                stl_Stok_kodu = tüketilecekUrunKodu,
                                stl_Satici_Kodu = "",
                                stl_miktari = tüketilecekUrunMiktari,
                                stl_teslim_miktari = 0,
                                stl_aciklama = üretilecekUrunKodu,
                                stl_aciklama2 = "",
                                stl_depo_no = 70150,
                                stl_kapat_fl = 0,
                                stl_projekodu = projeKodu,
                                stl_parti_kodu = "",
                                stl_lot_no = "0",
                                stl_OnaylayanKulNo = "0",
                                stl_cagrilabilir_fl = 1,
                                stl_harekettipi = 0,
                                stl_talep_eden = textBox.Text,
                                stl_kapatmanedenkod = "",
                                stl_KaynakSip_uid = Guid.Empty,
                                stl_HareketGrupKodu1 = isEmriKodu,
                                stl_HareketGrupKodu2 = "",
                                stl_HareketGrupKodu3 = "",
                                stl_birim_pntr = 1,
                                stl_Franchise_Fiyati = 0,
                            };
                            sayac++;
                            satirNo = sayac;
                            talepListesiComponent.Add(talep4);
                        }

                    }
                }
                // "DIN" ile baþlayan öðelerin toplamýný hesapla ve tek bir öðe olarak sakla
                var stokMiktariToplamlari = talepListesiComponent
                    .Where(talep => talep.stl_Stok_kodu.StartsWith("DIN"))
                    .GroupBy(talep => talep.stl_Stok_kodu)
                    .ToDictionary(
                        group => group.Key,
                        group => group.Sum(talep => talep.stl_miktari)
                    );

                // "DIN" ile baþlayan öðeleri güncelle
                foreach (var talep in talepListesiComponent.Where(talep => talep.stl_Stok_kodu.StartsWith("DIN")).ToList())
                {
                    talep.stl_miktari = stokMiktariToplamlari[talep.stl_Stok_kodu];
                }

                List<SATINALMA_TALEPLERI> yeniListe = new List<SATINALMA_TALEPLERI>();

                foreach (var talep in talepListesiComponent)
                {
                    bool eklendi = false;
                    if (talep.stl_Stok_kodu.StartsWith("DIN"))
                    {
                        // Yeni listeyi kontrol et, ayný kod zaten eklenmiþ mi?
                        foreach (var eleman in yeniListe)
                        {
                            if (eleman.stl_Stok_kodu == talep.stl_Stok_kodu)
                            {
                                eklendi = true;
                                break;
                            }
                        }

                        // Eðer eklenmediyse, yeni listeye ekle
                        if (!eklendi)
                        {
                            yeniListe.Add(talep);
                        }
                    }
                }
                List<SATINALMA_TALEPLERI> dinIleBaslamayanlar = new List<SATINALMA_TALEPLERI>();

                foreach (var talep in talepListesiComponent)
                {
                    if (!talep.stl_Stok_kodu.StartsWith("DIN"))
                    {
                        dinIleBaslamayanlar.Add(talep);
                    }
                }
                List<SATINALMA_TALEPLERI> birlesikListe = new List<SATINALMA_TALEPLERI>();

                birlesikListe.AddRange(yeniListe);
                birlesikListe.AddRange(dinIleBaslamayanlar);
                dbContext.SATINALMA_TALEPLERI.AddRange(talepListesiHamMadde);
                dbContext.SATINALMA_TALEPLERI.AddRange(talepListesiFason);
                dbContext.SATINALMA_TALEPLERI.AddRange(talepListesiResimliSA);
                dbContext.SATINALMA_TALEPLERI.AddRange(birlesikListe);
                dbContext.SaveChanges();

                MessageBox.Show("Veriler baþarýyla eklendi.");
            }



            else if (textBox1.Text.StartsWith("MX") || textBox1.Text.StartsWith("MS"))
            {
                List<SATINALMA_TALEPLERI> talepListesiHamMadde = new List<SATINALMA_TALEPLERI>();
                List<SATINALMA_TALEPLERI> talepListesiFason = new List<SATINALMA_TALEPLERI>();
                List<SATINALMA_TALEPLERI> talepListesiComponent = new List<SATINALMA_TALEPLERI>();
                List<SATINALMA_TALEPLERI> talepListesiResimliSA = new List<SATINALMA_TALEPLERI>();

                int nextStlEvrakSira = dbContext.SATINALMA_TALEPLERI.Max(item => item.stl_evrak_sira);
                int evrakSira = nextStlEvrakSira + 1; // Benzersiz bir þekilde ayarlanmalýdýr
                                                      // DataGridView'deki her bir satýr için
                int sayac = 0;
                int satirNo = sayac;
                foreach (DataGridViewRow row in advancedDataGridView1.Rows)
                {
                    if (!row.IsNewRow)
                    {

                        // DataGridView'den verileri al
                        string projeKodu = row.Cells["PROJE_KOD"].Value.ToString();
                        string isEmriKodu = row.Cells["ÝÞ_EMRÝ_KODU"].Value.ToString();
                        string üretilecekUrunKodu = row.Cells["ÜRETÝLECEK_ÜRÜN_KODU"].Value.ToString();
                        double üretilecekUrunMiktari = double.Parse(row.Cells["ÜRETÝLECEK_ÜRÜN_MÝKTARI"].Value.ToString());
                        string tüketilecekUrunKodu = row.Cells["TÜKETÝLECEK_ÜRÜN_KODU"].Value.ToString();
                        double tüketilecekUrunMiktari = double.Parse(row.Cells["TÜKETÝLECEK_ÜRÜN_MÝKTARI"].Value.ToString());
                        string tüketilecekUrunKoduAciklama = row.Cells["TÜKETÝLECEK_ÜRÜN_KODU_AÇIKLAMA"].Value.ToString();
                        string isemritipi = row.Cells["ÝSEMRÝ_TÝPÝ"].Value.ToString();
                        // Benzersiz deðerleri ayarlayýn
                        string evrakSeri = "T"; // Örnek olarak "T" olarak ayarlandý, siz iþ mantýðýnýza göre deðiþtirmelisiniz
                        string[] allowedEndings = new string[] { ".120", ".121", ".122", ".127" };
                        string[] allowedEndings1 = new string[] { ".000" };
                        string[] allowedEndings2 = new string[] { ".128" };

                        Dictionary<string, string> numaraKarsilik = new Dictionary<string, string>
                        {
                                { ".120", "LAZER KESÝM" },
                                { ".121", "OKSÝJEN KESÝM" },
                                { ".122", "SU JETÝ KESÝM" },
                                { ".127", "ROUTER KESÝM" }
                        };


                        var rsa1 = dbContext.URETIM_MALZEME_PLANLAMA
                            .Join(
                                dbContext.ISEMIRLERI,
                                u => u.upl_isemri,
                                i => i.is_Kod,
                                (u, i) => new { u, i }
                            )
                            .Where(x => x.u.upl_kodu == üretilecekUrunKodu &&
                                        x.i.is_ProjeKodu == projeKodu &&
                                        x.i.is_Kod == isEmriKodu &&
                                        string.IsNullOrEmpty(x.u.upl_urstokkod)) // Boþ veya null kontrolü
                            .Select(x => x.i.is_BagliOlduguIsemri)
                            .FirstOrDefault();

                        var rsa = dbContext.URETIM_MALZEME_PLANLAMA.Where(x => x.upl_isemri == rsa1 && !string.IsNullOrEmpty(x.upl_urstokkod)).Select(x => x.upl_urstokkod).FirstOrDefault();

                        var malzemeTanýmlar = customerList
                            .Where(x =>
                                x.ContainsKey("Part Number") &&
                                x["Part Number"].ToString() == tüketilecekUrunKodu &&
                                x.ContainsKey("Fason") &&
                                x["Fason"].ToString() != "EVET" &&
                                x.ContainsKey("Malzeme Taným") &&
                                !string.IsNullOrEmpty(x["Malzeme Taným"].ToString())
                            )
                            .Select(x => x["Malzeme Taným"].ToString())
                            .FirstOrDefault();




                        // KAYIÞLAR
                        if (Convert.ToBoolean(row.Cells["checkBoxColumn1"].Value) && malzemeTanýmlar != null)
                        {
                            SATINALMA_TALEPLERI talep5 = new SATINALMA_TALEPLERI
                            {
                                stl_DBCno = 0,
                                stl_SpecRECno = 0,
                                stl_iptal = 0,
                                stl_fileid = 118,
                                stl_hidden = 0,
                                stl_kilitli = 0,
                                stl_degisti = 0,
                                stl_checksum = 0,
                                stl_create_user = comboBox.Text,
                                stl_create_date = DateTime.Now,
                                stl_lastup_user = comboBox.Text,
                                stl_lastup_date = DateTime.Now,
                                stl_special1 = comboBox.Text,
                                stl_special2 = "",
                                stl_special3 = "",
                                stl_firmano = 0,
                                stl_subeno = 0,
                                stl_tarihi = selectedDateEvrakComp,
                                stl_teslim_tarihi = selectedDateTesComp,
                                stl_evrak_seri = evrakSeri,
                                stl_evrak_sira = evrakSira + 1,
                                stl_satir_no = satirNo,
                                stl_belge_no = textBoxBelge.Text,
                                stl_belge_tarihi = DateTime.Now.Date,
                                stl_Sor_Merk = projeKodu,
                                stl_Stok_kodu = malzemeTanýmlar.Length > 25 ? malzemeTanýmlar.Substring(0, 25) : malzemeTanýmlar,
                                stl_Satici_Kodu = "",
                                stl_miktari = tüketilecekUrunMiktari,
                                stl_teslim_miktari = 0,
                                stl_aciklama = üretilecekUrunKodu,
                                stl_aciklama2 = "",
                                stl_depo_no = 70150,
                                stl_kapat_fl = 0,
                                stl_projekodu = projeKodu,
                                stl_parti_kodu = "",
                                stl_lot_no = "0",
                                stl_OnaylayanKulNo = "0",
                                stl_cagrilabilir_fl = 1,
                                stl_harekettipi = 0,
                                stl_talep_eden = textBox.Text,
                                stl_kapatmanedenkod = "",
                                stl_KaynakSip_uid = Guid.Empty,
                                stl_HareketGrupKodu1 = isEmriKodu,
                                stl_HareketGrupKodu2 = "",
                                stl_HareketGrupKodu3 = "",
                                stl_birim_pntr = 1,
                                stl_Franchise_Fiyati = 0,
                            };
                            sayac++;
                            satirNo = sayac;
                            talepListesiResimliSA.Add(talep5);
                        }
                        //HAMMADDE-KESÝMLER
                        else if (Convert.ToBoolean(row.Cells["checkBoxColumn1"].Value) && allowedEndings.Any(end => üretilecekUrunKodu.EndsWith(end)))
                        {
                            string numara = allowedEndings.FirstOrDefault(end => üretilecekUrunKodu.EndsWith(end));
                            if (numara != null && numaraKarsilik.ContainsKey(numara))
                            {
                                string aciklama = üretilecekUrunMiktari + " Adet /" + numaraKarsilik[numara];

                                SATINALMA_TALEPLERI talep1 = new SATINALMA_TALEPLERI
                                {
                                    stl_DBCno = 0,
                                    stl_SpecRECno = 0,
                                    stl_iptal = 0,
                                    stl_fileid = 118,
                                    stl_hidden = 0,
                                    stl_kilitli = 0,
                                    stl_degisti = 0,
                                    stl_checksum = 0,
                                    stl_create_user = comboBox.Text,
                                    stl_create_date = DateTime.Now,
                                    stl_lastup_user = comboBox.Text,
                                    stl_lastup_date = DateTime.Now,
                                    stl_special1 = comboBox.Text,
                                    stl_special2 = "",
                                    stl_special3 = "",
                                    stl_firmano = 0,
                                    stl_subeno = 0,
                                    stl_tarihi = selectedDateEvrakHammade,
                                    stl_teslim_tarihi = selectedDateTesHammade,
                                    stl_evrak_seri = evrakSeri,
                                    stl_evrak_sira = evrakSira,
                                    stl_satir_no = satirNo,
                                    stl_belge_no = textBoxBelge.Text,
                                    stl_belge_tarihi = DateTime.Now.Date,
                                    stl_Sor_Merk = projeKodu,
                                    stl_Stok_kodu = tüketilecekUrunKodu,
                                    stl_Satici_Kodu = "",
                                    stl_miktari = tüketilecekUrunMiktari,
                                    stl_teslim_miktari = 0,
                                    stl_aciklama = aciklama,
                                    stl_aciklama2 = üretilecekUrunKodu,
                                    stl_depo_no = 70150,
                                    stl_kapat_fl = 0,
                                    stl_projekodu = projeKodu,
                                    stl_parti_kodu = "",
                                    stl_lot_no = "0",
                                    stl_OnaylayanKulNo = "0",
                                    stl_cagrilabilir_fl = 1,
                                    stl_harekettipi = 0,
                                    stl_talep_eden = textBox.Text,
                                    stl_kapatmanedenkod = "",
                                    stl_KaynakSip_uid = Guid.Empty,
                                    stl_HareketGrupKodu1 = isEmriKodu,
                                    stl_HareketGrupKodu2 = "",
                                    stl_HareketGrupKodu3 = "",
                                    stl_birim_pntr = 1,
                                    stl_Franchise_Fiyati = 0,
                                };
                                sayac++;
                                satirNo = sayac;
                                talepListesiHamMadde.Add(talep1);
                            }
                        }
                        //HAMMADDE-DEMÝRÇELÝK
                        else if (Convert.ToBoolean(row.Cells["checkBoxColumn1"].Value) && allowedEndings1.Any(end => üretilecekUrunKodu.EndsWith(end)))
                        {
                            string aciklama = tüketilecekUrunKoduAciklama + " / " + üretilecekUrunMiktari + " Adet";

                            SATINALMA_TALEPLERI talep1 = new SATINALMA_TALEPLERI
                            {
                                stl_DBCno = 0,
                                stl_SpecRECno = 0,
                                stl_iptal = 0,
                                stl_fileid = 118,
                                stl_hidden = 0,
                                stl_kilitli = 0,
                                stl_degisti = 0,
                                stl_checksum = 0,
                                stl_create_user = comboBox.Text,
                                stl_create_date = DateTime.Now,
                                stl_lastup_user = comboBox.Text,
                                stl_lastup_date = DateTime.Now,
                                stl_special1 = comboBox.Text,
                                stl_special2 = "",
                                stl_special3 = "",
                                stl_firmano = 0,
                                stl_subeno = 0,
                                stl_tarihi = selectedDateEvrakHammade,
                                stl_teslim_tarihi = selectedDateTesHammade,
                                stl_evrak_seri = evrakSeri,
                                stl_evrak_sira = evrakSira,
                                stl_satir_no = satirNo,
                                stl_belge_no = textBoxBelge.Text,
                                stl_belge_tarihi = DateTime.Now.Date,
                                stl_Sor_Merk = projeKodu,
                                stl_Stok_kodu = tüketilecekUrunKodu,
                                stl_Satici_Kodu = "",
                                stl_miktari = tüketilecekUrunMiktari,
                                stl_teslim_miktari = 0,
                                stl_aciklama = aciklama,
                                stl_aciklama2 = üretilecekUrunKodu,
                                stl_depo_no = 70150,
                                stl_kapat_fl = 0,
                                stl_projekodu = projeKodu,
                                stl_parti_kodu = "",
                                stl_lot_no = "0",
                                stl_OnaylayanKulNo = "0",
                                stl_cagrilabilir_fl = 1,
                                stl_harekettipi = 0,
                                stl_talep_eden = textBox.Text,
                                stl_kapatmanedenkod = "",
                                stl_KaynakSip_uid = Guid.Empty,
                                stl_HareketGrupKodu1 = isEmriKodu,
                                stl_HareketGrupKodu2 = "",
                                stl_HareketGrupKodu3 = "",
                                stl_birim_pntr = 1,
                                stl_Franchise_Fiyati = 0,
                            };
                            sayac++;
                            satirNo = sayac;
                            talepListesiHamMadde.Add(talep1);

                        }
                        //RESÝMLÝ SATIN ALMA
                        else if (Convert.ToBoolean(row.Cells["checkBoxColumn1"].Value) && allowedEndings2.Any(end => tüketilecekUrunKodu.EndsWith(end)) && malzemeTanýmlar == null)
                        {
                            SATINALMA_TALEPLERI talep3 = new SATINALMA_TALEPLERI
                            {
                                stl_DBCno = 0,
                                stl_SpecRECno = 0,
                                stl_iptal = 0,
                                stl_fileid = 118,
                                stl_hidden = 0,
                                stl_kilitli = 0,
                                stl_degisti = 0,
                                stl_checksum = 0,
                                stl_create_user = comboBox.Text,
                                stl_create_date = DateTime.Now,
                                stl_lastup_user = comboBox.Text,
                                stl_lastup_date = DateTime.Now,
                                stl_special1 = comboBox.Text,
                                stl_special2 = "",
                                stl_special3 = "",
                                stl_firmano = 0,
                                stl_subeno = 0,
                                stl_tarihi = selectedDateEvrakResSA,
                                stl_teslim_tarihi = selectedDateTesResSA,
                                stl_evrak_seri = evrakSeri,
                                stl_evrak_sira = evrakSira + 1,
                                stl_satir_no = satirNo,
                                stl_belge_no = textBoxBelge.Text,
                                stl_belge_tarihi = DateTime.Now.Date,
                                stl_Sor_Merk = projeKodu,
                                stl_Stok_kodu = üretilecekUrunKodu,
                                stl_Satici_Kodu = "",
                                stl_miktari = üretilecekUrunMiktari,
                                stl_teslim_miktari = 0,
                                stl_aciklama = rsa,
                                stl_aciklama2 = "",
                                stl_depo_no = 70150,
                                stl_kapat_fl = 0,
                                stl_projekodu = projeKodu,
                                stl_parti_kodu = "",
                                stl_lot_no = "0",
                                stl_OnaylayanKulNo = "0",
                                stl_cagrilabilir_fl = 1,
                                stl_harekettipi = 0,
                                stl_talep_eden = textBox.Text,
                                stl_kapatmanedenkod = "",
                                stl_KaynakSip_uid = Guid.Empty,
                                stl_HareketGrupKodu1 = isEmriKodu,
                                stl_HareketGrupKodu2 = "",
                                stl_HareketGrupKodu3 = "",
                                stl_birim_pntr = 1,
                                stl_Franchise_Fiyati = 0,
                            };
                            sayac++;
                            satirNo = sayac;
                            talepListesiResimliSA.Add(talep3);
                        }

                        // MONTAJ ÝÞ EMRÝ VE COMPANENT ÝÞ EMRÝ OLABÝLÝCEKLER.
                        else if (Convert.ToBoolean(row.Cells["checkBoxColumn1"].Value))
                        {
                            SATINALMA_TALEPLERI talep4 = new SATINALMA_TALEPLERI
                            {
                                stl_DBCno = 0,
                                stl_SpecRECno = 0,
                                stl_iptal = 0,
                                stl_fileid = 118,
                                stl_hidden = 0,
                                stl_kilitli = 0,
                                stl_degisti = 0,
                                stl_checksum = 0,
                                stl_create_user = comboBox.Text,
                                stl_create_date = DateTime.Now,
                                stl_lastup_user = comboBox.Text,
                                stl_lastup_date = DateTime.Now,
                                stl_special1 = comboBox.Text,
                                stl_special2 = "",
                                stl_special3 = "",
                                stl_firmano = 0,
                                stl_subeno = 0,
                                stl_tarihi = selectedDateEvrakComp,
                                stl_teslim_tarihi = selectedDateTesComp,
                                stl_evrak_seri = evrakSeri,
                                stl_evrak_sira = evrakSira + 2,
                                stl_satir_no = satirNo,
                                stl_belge_no = textBoxBelge.Text,
                                stl_belge_tarihi = DateTime.Now.Date,
                                stl_Sor_Merk = projeKodu,
                                stl_Stok_kodu = tüketilecekUrunKodu,
                                stl_Satici_Kodu = "",
                                stl_miktari = tüketilecekUrunMiktari,
                                stl_teslim_miktari = 0,
                                stl_aciklama = üretilecekUrunKodu,
                                stl_aciklama2 = "",
                                stl_depo_no = 70150,
                                stl_kapat_fl = 0,
                                stl_projekodu = projeKodu,
                                stl_parti_kodu = "",
                                stl_lot_no = "0",
                                stl_OnaylayanKulNo = "0",
                                stl_cagrilabilir_fl = 1,
                                stl_harekettipi = 0,
                                stl_talep_eden = textBox.Text,
                                stl_kapatmanedenkod = "",
                                stl_KaynakSip_uid = Guid.Empty,
                                stl_HareketGrupKodu1 = isEmriKodu,
                                stl_HareketGrupKodu2 = "",
                                stl_HareketGrupKodu3 = "",
                                stl_birim_pntr = 1,
                                stl_Franchise_Fiyati = 0,
                            };
                            sayac++;
                            satirNo = sayac;
                            talepListesiComponent.Add(talep4);
                        }




                    }
                }
                // "DIN" ile baþlayan öðelerin toplamýný hesapla ve tek bir öðe olarak sakla
                var stokMiktariToplamlari = talepListesiComponent
                    .Where(talep => talep.stl_Stok_kodu.StartsWith("DIN"))
                    .GroupBy(talep => talep.stl_Stok_kodu)
                    .ToDictionary(
                        group => group.Key,
                        group => group.Sum(talep => talep.stl_miktari)
                    );

                // "DIN" ile baþlayan öðeleri güncelle
                foreach (var talep in talepListesiComponent.Where(talep => talep.stl_Stok_kodu.StartsWith("DIN")).ToList())
                {
                    talep.stl_miktari = stokMiktariToplamlari[talep.stl_Stok_kodu];
                }

                List<SATINALMA_TALEPLERI> yeniListe = new List<SATINALMA_TALEPLERI>();

                foreach (var talep in talepListesiComponent)
                {
                    bool eklendi = false;
                    if (talep.stl_Stok_kodu.StartsWith("DIN"))
                    {
                        // Yeni listeyi kontrol et, ayný kod zaten eklenmiþ mi?
                        foreach (var eleman in yeniListe)
                        {
                            if (eleman.stl_Stok_kodu == talep.stl_Stok_kodu)
                            {
                                eklendi = true;
                                break;
                            }
                        }

                        // Eðer eklenmediyse, yeni listeye ekle
                        if (!eklendi)
                        {
                            yeniListe.Add(talep);
                        }
                    }
                }
                List<SATINALMA_TALEPLERI> dinIleBaslamayanlar = new List<SATINALMA_TALEPLERI>();

                foreach (var talep in talepListesiComponent)
                {
                    if (!talep.stl_Stok_kodu.StartsWith("DIN"))
                    {
                        dinIleBaslamayanlar.Add(talep);
                    }
                }
                List<SATINALMA_TALEPLERI> birlesikListe = new List<SATINALMA_TALEPLERI>();

                birlesikListe.AddRange(yeniListe);
                birlesikListe.AddRange(dinIleBaslamayanlar);
                dbContext.SATINALMA_TALEPLERI.AddRange(talepListesiHamMadde);
                dbContext.SATINALMA_TALEPLERI.AddRange(talepListesiFason);
                dbContext.SATINALMA_TALEPLERI.AddRange(talepListesiResimliSA);
                dbContext.SATINALMA_TALEPLERI.AddRange(birlesikListe);

                dbContext.SaveChanges();

                MessageBox.Show("Veriler baþarýyla eklendi.");

            }
            sebepSecimiForm.Close();
        };

            sebepSecimiForm.Controls.AddRange(new Control[] { comboBox, textBox, onayButton });

            sebepSecimiForm.ShowDialog();






        }
        private void button1_Click(object sender, EventArgs e)
        {
            siparisGiris();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            InsertDataToDatabase();
        }

        private void dataGridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex == advancedDataGridView1.Columns["TÜKETÝLECEK_ÜRÜN_KODU"].Index)
            {
                if (advancedDataGridView1.Columns[e.ColumnIndex].HeaderCell.SortGlyphDirection == SortOrder.Ascending)
                {
                    // Sýralamayý artan sýraya göre yap
                    advancedDataGridView1.Sort(advancedDataGridView1.Columns["TÜKETÝLECEK_ÜRÜN_KODU"], ListSortDirection.Ascending);
                }
                else
                {
                    // Sýralamayý azalan sýraya göre yap
                    advancedDataGridView1.Sort(advancedDataGridView1.Columns["TÜKETÝLECEK_ÜRÜN_KODU"], ListSortDirection.Descending);
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            List<byte> durumu = new List<byte>() { 0, 1, 2 };
            comboBox1.DataSource = durumu;


            toolStripLabel1.ToolTipText =

                "1.Ýþ Emri Kodu Gir\n" +

                "2.Tarih Seç\n" +

                "3.Ýþ Emri Durumunu Seç\n" +

                "4.Verileri Listele\n" +

                "5.Evrak Tarihlerini Seç\n" +

                "6.Teslim Tarihlerini Seç\n" +

                "7.Talep oluþtur";


        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                // Ana form kapatýldýðýnda tüm uygulamayý kapat
                Application.Exit();
            }
        }


        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            // DataGridView'deki verileri bir DataTable'a kopyalayýn
            DataTable dt = new DataTable();

            foreach (DataGridViewColumn column in advancedDataGridView1.Columns)
            {
                // Eðer ValueType null ise, varsayýlan bir veri türü kullanabilirsiniz.
                Type columnType = column.ValueType ?? typeof(string);
                dt.Columns.Add(column.HeaderText, columnType);
            }

            // Satýrlarý ekle
            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                DataRow dataRow = dt.NewRow();
                foreach (DataGridViewCell cell in row.Cells)
                {
                    dataRow[cell.ColumnIndex] = cell.Value;
                }
                dt.Rows.Add(dataRow);
            }

            // Excel uygulamasýný baþlatýn
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;

            // Yeni bir Excel çalýþma kitabý oluþturun
            Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];

            // DataTable'ý Excel çalýþma sayfasýna aktarýn (tablo baþlýklarýný da dahil etmek için)
            int rowIndex = 1;

            // Baþlýklarý yaz
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



        private void SetPentagonRegion(Button button)
        {
            GraphicsPath path = new GraphicsPath();

            // Beþgenin köþe noktalarýný belirleyin
            Point[] pentagonPoints = new Point[]
            {
        new Point(button.Width / 2, 0),
        new Point(button.Width, button.Height / 3),
        new Point(button.Width * 2 / 3, button.Height),
        new Point(button.Width / 3, button.Height),
        new Point(0, button.Height / 3)
            };

            // Beþgeni path'e ekleyin
            path.AddPolygon(pentagonPoints);

            // Path'i kullanarak yeni bir bölge oluþturun
            button.Region = new Region(path);
        }

        private void fASONÝMALATTALEBÝToolStripMenuItem_Click(object sender, EventArgs e)
        {
            EXCEL_TALEP et = new EXCEL_TALEP();
            et.Show();
        }

        private void pROFORMAToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ProformaSiparis ps = new ProformaSiparis();
            ps.Show();
        }

        private void oTOMASYONToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Otomasyon oto = new Otomasyon();
            oto.Show();
        }

        private void eSKÝKODToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TALEB_MALÝYET TM = new TALEB_MALÝYET();
            TM.Show();
        }

        private void yENÝKODToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TALEP_MALÝYET_YENÝKOD_ TMYK = new TALEP_MALÝYET_YENÝKOD_();
            TMYK.Show();
        }

        private void bOMDAROTADÜZELTMEToolStripMenuItem_Click(object sender, EventArgs e)
        {
            BOM_ROTA BR = new BOM_ROTA();
            BR.Show();
        }

        private void vERÝLENTEKLÝFMALÝYETHESABIToolStripMenuItem_Click(object sender, EventArgs e)
        {
            VERÝLEN_TEKLÝF_MALÝYET_HESABI VMH = new VERÝLEN_TEKLÝF_MALÝYET_HESABI();
            VMH.Show();
        }

        private void kODBULMAVEMALÝYETToolStripMenuItem_Click(object sender, EventArgs e)
        {
            KOD_BULMA KD = new KOD_BULMA();
            KD.Show();
        }

        private void sÜRETRANSFERÝToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SURE_TRANSFERI ST = new SURE_TRANSFERI();
            ST.Show();
        }

        private void mARKAKODUATAMAToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MARKA_KODU MK = new MARKA_KODU();
            MK.Show();
        }

        DataTable data = new DataTable();
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog file = new OpenFileDialog();
                file.Filter = "Excel Dosyasý |*.xlsx";
                file.ShowDialog();

                string tamYol = file.FileName;

                string baglantiAdresi = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + tamYol + ";Extended Properties='Excel 12.0;IMEX=1;'";

                using (OleDbConnection baglanti = new OleDbConnection(baglantiAdresi))
                {
                    OleDbCommand komut = new OleDbCommand("Select * From [" + textBox2.Text + "$]", baglanti);

                    baglanti.Open();

                    OleDbDataAdapter da = new OleDbDataAdapter(komut);
                    DataTable data = new DataTable();
                    da.Fill(data);

                    // Checkbox sütununu eklemiyoruz

                    // Verileri bir listeye aktar

                    foreach (DataRow row in data.Rows)
                    {
                        Dictionary<string, object> item = new Dictionary<string, object>();
                        foreach (DataColumn column in data.Columns)
                        {
                            item[column.ColumnName] = row[column]; // Sütun adýný anahtar, hücre deðerini deðer olarak ekle
                        }
                        customerList.Add(item);
                    }

                    MessageBox.Show("BOM Listesi Yüklendi");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

        private void bOMERPMRPKONTROLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            BOM_KONTROL BK = new BOM_KONTROL();
            BK.Show();
        }

        private void tEKLÝFVERMEToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TEKLÝF TF = new TEKLÝF();
            TF.Show();
        }

        private void lEADTÝMEToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AÐAC_TERMÝN AT = new AÐAC_TERMÝN();
            AT.Show();
        }

        private void tALEPFÝYATLÝSTESÝToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TALEP_FÝYAT TF = new TALEP_FÝYAT();
            TF.Show();
        }

        private void kALANSÝPARÝÞDETAYIToolStripMenuItem_Click(object sender, EventArgs e)
        {
            KALAN_SÝPARÝS KS = new KALAN_SÝPARÝS();
            KS.Show();
        }

        private void rOTAVEÝÞMERKEZÝÝZLEMEToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ROTA_ÝSMERKEZÝ RI = new ROTA_ÝSMERKEZÝ();
            RI.Show();
        }

        private void toolStripTextBox1_Click(object sender, EventArgs e)
        {
            YÖNETÝM_EKRANI YE = new YÖNETÝM_EKRANI();
            YE.Show();
        }

        private void mRPÖNCESÝTAHMÝNÝSÜREToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MRP_ONCESÝ_TAHMÝNÝ_SURE MRPOTS = new MRP_ONCESÝ_TAHMÝNÝ_SURE();
            MRPOTS.Show();
        }
    }
}