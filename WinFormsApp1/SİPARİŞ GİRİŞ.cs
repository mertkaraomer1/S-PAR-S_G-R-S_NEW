using Castle.Core.Resource;
using S�PAR��_G�R��;
using S�PAR��_G�R��.Tables;
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
            table.Columns.Add("��_EMR�_KODU", typeof(string));
            table.Columns.Add("S�PAR��_SER�", typeof(string));
            table.Columns.Add("S�PAR��_NO", typeof(string));
            table.Columns.Add("�RET�LECEK_�R�N_KODU", typeof(string));
            table.Columns.Add("�RET�LECEK_�R�N_ADI", typeof(string));
            table.Columns.Add("�RET�LECEK_�R�N_M�KTARI", typeof(double));
            table.Columns.Add("T�KET�LECEK_�R�N_KODU_A�IKLAMA", typeof(string));
            table.Columns.Add("T�KET�LECEK_�R�N_KODU", typeof(string));
            table.Columns.Add("T�KET�LECEK_�R�N_ADI", typeof(string));
            table.Columns.Add("T�KET�LECEK_�R�N_M�KTARI", typeof(double));
            table.Columns.Add("�SEMR�_T�P�", typeof(string));
            table.Columns.Add("DURUMU", typeof(byte));
            table.Columns.Add("TAR�H", typeof(DateTime));


            var sorgu = from isEmir in dbContext.ISEMIRLERI
                        join uretim in dbContext.URETIM_MALZEME_PLANLAMA on isEmir.is_Kod equals uretim.upl_isemri
                        where isEmir.is_Kod.Substring(0, 8) == textBox1.Text &&
                              isEmir.is_create_date.Date >= selectedDate &&
                              isEmir.is_create_date.Date <= selectedDate1 &&
                              isEmir.is_EmriDurumu == Durum && // Durum kontrol�
                              dbContext.ISEMIRLERI_USER.Any(u => u.Record_uid == isEmir.is_Guid && u.is_emri_tipi != null) && // Null kontrol�
                              uretim.upl_satirno != 0
                        orderby uretim.upl_kodu ascending
                        select new
                        {
                            SATIR_NO = uretim.upl_satirno,
                            PROJE_KOD = isEmir.is_ProjeKodu,
                            ��_EMR�_KODU = isEmir.is_Kod,
                            S�PAR��_SER� = isEmir.is_SiparisNo_Seri,
                            S�PAR��_NO = isEmir.is_SiparisNo_Sira,
                            �RET�LECEK_�R�N_KODU = uretim.upl_urstokkod,
                            �RET�LECEK_�R�N_ADI = dbContext.STOKLAR.FirstOrDefault(u => u.sto_kod == uretim.upl_urstokkod).sto_isim,
                            �RET�LECEK_�R�N_M�KTARI = uretim.upl_uret_miktar,
                            T�KET�LECEK_�R�N_KODU_A�IKLAMA = uretim.upl_aciklama,
                            T�KET�LECEK_�R�N_KODU = uretim.upl_kodu,
                            T�KET�LECEK_�R�N_ADI = dbContext.STOKLAR.FirstOrDefault(u => u.sto_kod == uretim.upl_kodu).sto_isim,
                            T�KET�LECEK_�R�N_M�KTARI = uretim.upl_sarfmiktari,
                            �SEMR�_T�P� = dbContext.ISEMIRLERI_USER.FirstOrDefault(u => u.Record_uid == isEmir.is_Guid).is_emri_tipi,
                            DURUMU = isEmir.is_EmriDurumu,
                            TAR�H = isEmir.is_create_date.Date
                        };




            // Sorgu sonucundaki verileri liste olarak al�n
            var veriListesi = sorgu.ToList();



            // E�er veri yoksa kullan�c�ya mesaj g�sterin ve DataGridView'i temizleyin
            if (veriListesi.Count == 0)
            {
                MessageBox.Show("Se�ilen tarihte veri bulunmamaktad�r.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                advancedDataGridView1.DataSource = null; // DataGridView'i temizle
            }
            else
            {
                // DataGridView ba�l�k sat�r�n�n g�r�n�m�n� ayarlay�n
                advancedDataGridView1.EnableHeadersVisualStyles = false;
                advancedDataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.DimGray;
                advancedDataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

                // �zel bir yaz� tipi tan�mlay�n
                Font baslikYaziTipi = new Font("Arial", 10, FontStyle.Regular);

                // DataGridView ba�l�k sat�r�n�n yaz� tipini ayarlay�n
                advancedDataGridView1.ColumnHeadersDefaultCellStyle.Font = baslikYaziTipi;
                foreach (var urun in veriListesi)
                {
                    table.Rows.Add(urun.SATIR_NO, urun.PROJE_KOD, urun.��_EMR�_KODU, urun.S�PAR��_SER�, urun.S�PAR��_NO, urun.�RET�LECEK_�R�N_KODU, urun.�RET�LECEK_�R�N_ADI, urun.�RET�LECEK_�R�N_M�KTARI,
                        urun.T�KET�LECEK_�R�N_KODU_A�IKLAMA, urun.T�KET�LECEK_�R�N_KODU, urun.T�KET�LECEK_�R�N_ADI, urun.T�KET�LECEK_�R�N_M�KTARI, urun.�SEMR�_T�P�, urun.DURUMU, urun.TAR�H);
                }

                // DataGridView'e verileri y�kleyin
                advancedDataGridView1.DataSource = table;
                // DataGridView'e verileri y�kledikten sonra CheckBox'lar� i�aretle
                for (int rowIndex = 0; rowIndex < advancedDataGridView1.Rows.Count; rowIndex++)
                {
                    var veri = veriListesi[rowIndex]; // Mevcut sat�r�n verisi
                    bool shouldCheck = false;

                    // �lk ko�ullar�n kontrol�
                    if ((veri.�SEMR�_T�P� == "SATINALMA_IE" &&
                         !new[] { ".128" }.Any(prefix => veri.�RET�LECEK_�R�N_KODU.EndsWith(prefix))) ||
                        (veri.�SEMR�_T�P� == "KK_IE" &&
                         new[] { ".128" }.Any(prefix => veri.T�KET�LECEK_�R�N_KODU.EndsWith(prefix))) ||
                        (veri.�SEMR�_T�P� == "MONTAJ_IE" &&
                         (!dbContext.URUNLER.Any(x => x.uru_stok_kod == veri.T�KET�LECEK_�R�N_KODU)))
                    )
                    {
                        shouldCheck = true;
                    }

                    if (veri.�SEMR�_T�P� == "MONTAJ_IE")
                    {
                        foreach (var x in customerList)
                        {
                            if (
                                x.ContainsKey("Part Number") &&
                                new[] { "02" }.Any(prefix => veri.T�KET�LECEK_�R�N_KODU.StartsWith(prefix)) &&
                                x["Part Number"].ToString() == veri.T�KET�LECEK_�R�N_KODU &&
                                x.ContainsKey("Fason") &&
                                x["Fason"].ToString() != "EVET" &&
                                x.ContainsKey("Malzeme Tan�m") &&
                                !string.IsNullOrEmpty(x["Malzeme Tan�m"].ToString()) &&
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


                    // Burada shouldCheck de�i�kenini kullanarak gerekli i�lemleri yapabilirsiniz


                    // E�er ko�ullar sa�lan�yorsa checkbox'� i�aretle
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
                Text = "Ara Verme Sebebinizi Se�iniz...",
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
                Visible = false  // Ba�lang��ta g�r�nmez
            };
            TextBox textBoxBelge = new TextBox
            {
                Location = new Point(50, 100),
                Width = 200,
                Visible = false  // Ba�lang��ta g�r�nmez
            };
            Button onayButton = new Button
            {
                Text = "Onay",
                Location = new Point(100, 100),
                Size = new Size(60, 30),
            };
            // ComboBox'�n SelectedIndexChanged etkinli�i
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
                    textBoxBelge.Text = "S�";
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
                    textBoxBelge.Text = "�B";
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
                int evrakSira = nextStlEvrakSira + 1; // Benzersiz bir �ekilde ayarlanmal�d�r
                                                      // DataGridView'deki her bir sat�r i�in
                int sayac = 0;
                int satirNo = sayac;
                foreach (DataGridViewRow row in advancedDataGridView1.Rows)
                {


                    if (!row.IsNewRow)
                    {

                        // DataGridView'den verileri al
                        string projeKodu = row.Cells["PROJE_KOD"].Value.ToString();
                        string isEmriKodu = row.Cells["��_EMR�_KODU"].Value.ToString();
                        string �retilecekUrunKodu = row.Cells["�RET�LECEK_�R�N_KODU"].Value.ToString();
                        double �retilecekUrunMiktari = double.Parse(row.Cells["�RET�LECEK_�R�N_M�KTARI"].Value.ToString());
                        string t�ketilecekUrunKodu = row.Cells["T�KET�LECEK_�R�N_KODU"].Value.ToString();
                        double t�ketilecekUrunMiktari = double.Parse(row.Cells["T�KET�LECEK_�R�N_M�KTARI"].Value.ToString());
                        string t�ketilecekUrunKoduAciklama = row.Cells["T�KET�LECEK_�R�N_KODU_A�IKLAMA"].Value.ToString();
                        string isemritipi = row.Cells["�SEMR�_T�P�"].Value.ToString();
                        string siparisseri = row.Cells["S�PAR��_SER�"].Value.ToString();
                        string sipariss�ra = row.Cells["S�PAR��_NO"].Value.ToString();
                        // Benzersiz de�erleri ayarlay�n
                        string evrakSeri = "T"; // �rnek olarak "T" olarak ayarland�, siz i� mant���n�za g�re de�i�tirmelisiniz

                        string SipSeriS�ra = siparisseri + "-" + sipariss�ra;
                        string[] allowedEndings = new string[] { ".120", ".121", ".122", ".127" };
                        string[] allowedEndings1 = new string[] { ".000" };
                        string[] allowedEndings2 = new string[] { ".128" };
                        Dictionary<string, string> numaraKarsilik = new Dictionary<string, string>
                        {
                                { ".120", "LAZER KES�M" },
                                { ".121", "OKS�JEN KES�M" },
                                { ".122", "SU JET� KES�M" },
                                { ".127", "ROUTER KES�M" }
                        };
                        var rsa = dbContext.URETIM_MALZEME_PLANLAMA.Where(x => x.upl_kodu == �retilecekUrunKodu).Select(x => x.upl_urstokkod).FirstOrDefault();
                        //HAMMADDE-KES�MLER
                        if (Convert.ToBoolean(row.Cells["checkBoxColumn1"].Value) && allowedEndings.Any(end => �retilecekUrunKodu.EndsWith(end)))
                        {
                            string numara = allowedEndings.FirstOrDefault(end => �retilecekUrunKodu.EndsWith(end));
                            if (numara != null && numaraKarsilik.ContainsKey(numara))
                            {
                                string aciklama = �retilecekUrunMiktari + " Adet /" + numaraKarsilik[numara];

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
                                    stl_Sor_Merk = SipSeriS�ra,
                                    stl_Stok_kodu = t�ketilecekUrunKodu,
                                    stl_Satici_Kodu = "",
                                    stl_miktari = t�ketilecekUrunMiktari,
                                    stl_teslim_miktari = 0,
                                    stl_aciklama = aciklama + "&" + textBox1.Text,
                                    stl_aciklama2 = �retilecekUrunKodu,
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
                        //HAMMADDE-DEM�R�EL�K
                        else if (Convert.ToBoolean(row.Cells["checkBoxColumn1"].Value) && allowedEndings1.Any(end => �retilecekUrunKodu.EndsWith(end)))
                        {
                            string aciklama = t�ketilecekUrunKoduAciklama + " / " + �retilecekUrunMiktari + " Adet";

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
                                stl_Stok_kodu = t�ketilecekUrunKodu,
                                stl_Satici_Kodu = "",
                                stl_miktari = t�ketilecekUrunMiktari,
                                stl_teslim_miktari = 0,
                                stl_aciklama = aciklama + "&" + textBox1.Text,
                                stl_aciklama2 = �retilecekUrunKodu,
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
                        //RES�ML� SATIN ALMA
                        else if (Convert.ToBoolean(row.Cells["checkBoxColumn1"].Value) && allowedEndings2.Any(end => t�ketilecekUrunKodu.EndsWith(end)))
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
                                stl_Sor_Merk = SipSeriS�ra,
                                stl_Stok_kodu = �retilecekUrunKodu,
                                stl_Satici_Kodu = "",
                                stl_miktari = �retilecekUrunMiktari,
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
                        // MONTAJ �� EMR� VE COMPANENT �� EMR� OLAB�L�CEKLER.
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
                                stl_Sor_Merk = SipSeriS�ra,
                                stl_Stok_kodu = t�ketilecekUrunKodu,
                                stl_Satici_Kodu = "",
                                stl_miktari = t�ketilecekUrunMiktari,
                                stl_teslim_miktari = 0,
                                stl_aciklama = �retilecekUrunKodu,
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
                // "DIN" ile ba�layan ��elerin toplam�n� hesapla ve tek bir ��e olarak sakla
                var stokMiktariToplamlari = talepListesiComponent
                    .Where(talep => talep.stl_Stok_kodu.StartsWith("DIN"))
                    .GroupBy(talep => talep.stl_Stok_kodu)
                    .ToDictionary(
                        group => group.Key,
                        group => group.Sum(talep => talep.stl_miktari)
                    );

                // "DIN" ile ba�layan ��eleri g�ncelle
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
                        // Yeni listeyi kontrol et, ayn� kod zaten eklenmi� mi?
                        foreach (var eleman in yeniListe)
                        {
                            if (eleman.stl_Stok_kodu == talep.stl_Stok_kodu)
                            {
                                eklendi = true;
                                break;
                            }
                        }

                        // E�er eklenmediyse, yeni listeye ekle
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

                MessageBox.Show("Veriler ba�ar�yla eklendi.");
            }



            else if (textBox1.Text.StartsWith("MX") || textBox1.Text.StartsWith("MS"))
            {
                List<SATINALMA_TALEPLERI> talepListesiHamMadde = new List<SATINALMA_TALEPLERI>();
                List<SATINALMA_TALEPLERI> talepListesiFason = new List<SATINALMA_TALEPLERI>();
                List<SATINALMA_TALEPLERI> talepListesiComponent = new List<SATINALMA_TALEPLERI>();
                List<SATINALMA_TALEPLERI> talepListesiResimliSA = new List<SATINALMA_TALEPLERI>();

                int nextStlEvrakSira = dbContext.SATINALMA_TALEPLERI.Max(item => item.stl_evrak_sira);
                int evrakSira = nextStlEvrakSira + 1; // Benzersiz bir �ekilde ayarlanmal�d�r
                                                      // DataGridView'deki her bir sat�r i�in
                int sayac = 0;
                int satirNo = sayac;
                foreach (DataGridViewRow row in advancedDataGridView1.Rows)
                {
                    if (!row.IsNewRow)
                    {

                        // DataGridView'den verileri al
                        string projeKodu = row.Cells["PROJE_KOD"].Value.ToString();
                        string isEmriKodu = row.Cells["��_EMR�_KODU"].Value.ToString();
                        string �retilecekUrunKodu = row.Cells["�RET�LECEK_�R�N_KODU"].Value.ToString();
                        double �retilecekUrunMiktari = double.Parse(row.Cells["�RET�LECEK_�R�N_M�KTARI"].Value.ToString());
                        string t�ketilecekUrunKodu = row.Cells["T�KET�LECEK_�R�N_KODU"].Value.ToString();
                        double t�ketilecekUrunMiktari = double.Parse(row.Cells["T�KET�LECEK_�R�N_M�KTARI"].Value.ToString());
                        string t�ketilecekUrunKoduAciklama = row.Cells["T�KET�LECEK_�R�N_KODU_A�IKLAMA"].Value.ToString();
                        string isemritipi = row.Cells["�SEMR�_T�P�"].Value.ToString();
                        // Benzersiz de�erleri ayarlay�n
                        string evrakSeri = "T"; // �rnek olarak "T" olarak ayarland�, siz i� mant���n�za g�re de�i�tirmelisiniz
                        string[] allowedEndings = new string[] { ".120", ".121", ".122", ".127" };
                        string[] allowedEndings1 = new string[] { ".000" };
                        string[] allowedEndings2 = new string[] { ".128" };

                        Dictionary<string, string> numaraKarsilik = new Dictionary<string, string>
                        {
                                { ".120", "LAZER KES�M" },
                                { ".121", "OKS�JEN KES�M" },
                                { ".122", "SU JET� KES�M" },
                                { ".127", "ROUTER KES�M" }
                        };


                        var rsa1 = dbContext.URETIM_MALZEME_PLANLAMA
                            .Join(
                                dbContext.ISEMIRLERI,
                                u => u.upl_isemri,
                                i => i.is_Kod,
                                (u, i) => new { u, i }
                            )
                            .Where(x => x.u.upl_kodu == �retilecekUrunKodu &&
                                        x.i.is_ProjeKodu == projeKodu &&
                                        x.i.is_Kod == isEmriKodu &&
                                        string.IsNullOrEmpty(x.u.upl_urstokkod)) // Bo� veya null kontrol�
                            .Select(x => x.i.is_BagliOlduguIsemri)
                            .FirstOrDefault();

                        var rsa = dbContext.URETIM_MALZEME_PLANLAMA.Where(x => x.upl_isemri == rsa1 && !string.IsNullOrEmpty(x.upl_urstokkod)).Select(x => x.upl_urstokkod).FirstOrDefault();

                        var malzemeTan�mlar = customerList
                            .Where(x =>
                                x.ContainsKey("Part Number") &&
                                x["Part Number"].ToString() == t�ketilecekUrunKodu &&
                                x.ContainsKey("Fason") &&
                                x["Fason"].ToString() != "EVET" &&
                                x.ContainsKey("Malzeme Tan�m") &&
                                !string.IsNullOrEmpty(x["Malzeme Tan�m"].ToString())
                            )
                            .Select(x => x["Malzeme Tan�m"].ToString())
                            .FirstOrDefault();




                        // KAYI�LAR
                        if (Convert.ToBoolean(row.Cells["checkBoxColumn1"].Value) && malzemeTan�mlar != null)
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
                                stl_Stok_kodu = malzemeTan�mlar.Length > 25 ? malzemeTan�mlar.Substring(0, 25) : malzemeTan�mlar,
                                stl_Satici_Kodu = "",
                                stl_miktari = t�ketilecekUrunMiktari,
                                stl_teslim_miktari = 0,
                                stl_aciklama = �retilecekUrunKodu,
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
                        //HAMMADDE-KES�MLER
                        else if (Convert.ToBoolean(row.Cells["checkBoxColumn1"].Value) && allowedEndings.Any(end => �retilecekUrunKodu.EndsWith(end)))
                        {
                            string numara = allowedEndings.FirstOrDefault(end => �retilecekUrunKodu.EndsWith(end));
                            if (numara != null && numaraKarsilik.ContainsKey(numara))
                            {
                                string aciklama = �retilecekUrunMiktari + " Adet /" + numaraKarsilik[numara];

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
                                    stl_Stok_kodu = t�ketilecekUrunKodu,
                                    stl_Satici_Kodu = "",
                                    stl_miktari = t�ketilecekUrunMiktari,
                                    stl_teslim_miktari = 0,
                                    stl_aciklama = aciklama,
                                    stl_aciklama2 = �retilecekUrunKodu,
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
                        //HAMMADDE-DEM�R�EL�K
                        else if (Convert.ToBoolean(row.Cells["checkBoxColumn1"].Value) && allowedEndings1.Any(end => �retilecekUrunKodu.EndsWith(end)))
                        {
                            string aciklama = t�ketilecekUrunKoduAciklama + " / " + �retilecekUrunMiktari + " Adet";

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
                                stl_Stok_kodu = t�ketilecekUrunKodu,
                                stl_Satici_Kodu = "",
                                stl_miktari = t�ketilecekUrunMiktari,
                                stl_teslim_miktari = 0,
                                stl_aciklama = aciklama,
                                stl_aciklama2 = �retilecekUrunKodu,
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
                        //RES�ML� SATIN ALMA
                        else if (Convert.ToBoolean(row.Cells["checkBoxColumn1"].Value) && allowedEndings2.Any(end => t�ketilecekUrunKodu.EndsWith(end)) && malzemeTan�mlar == null)
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
                                stl_Stok_kodu = �retilecekUrunKodu,
                                stl_Satici_Kodu = "",
                                stl_miktari = �retilecekUrunMiktari,
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

                        // MONTAJ �� EMR� VE COMPANENT �� EMR� OLAB�L�CEKLER.
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
                                stl_Stok_kodu = t�ketilecekUrunKodu,
                                stl_Satici_Kodu = "",
                                stl_miktari = t�ketilecekUrunMiktari,
                                stl_teslim_miktari = 0,
                                stl_aciklama = �retilecekUrunKodu,
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
                // "DIN" ile ba�layan ��elerin toplam�n� hesapla ve tek bir ��e olarak sakla
                var stokMiktariToplamlari = talepListesiComponent
                    .Where(talep => talep.stl_Stok_kodu.StartsWith("DIN"))
                    .GroupBy(talep => talep.stl_Stok_kodu)
                    .ToDictionary(
                        group => group.Key,
                        group => group.Sum(talep => talep.stl_miktari)
                    );

                // "DIN" ile ba�layan ��eleri g�ncelle
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
                        // Yeni listeyi kontrol et, ayn� kod zaten eklenmi� mi?
                        foreach (var eleman in yeniListe)
                        {
                            if (eleman.stl_Stok_kodu == talep.stl_Stok_kodu)
                            {
                                eklendi = true;
                                break;
                            }
                        }

                        // E�er eklenmediyse, yeni listeye ekle
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

                MessageBox.Show("Veriler ba�ar�yla eklendi.");

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
            if (e.ColumnIndex == advancedDataGridView1.Columns["T�KET�LECEK_�R�N_KODU"].Index)
            {
                if (advancedDataGridView1.Columns[e.ColumnIndex].HeaderCell.SortGlyphDirection == SortOrder.Ascending)
                {
                    // S�ralamay� artan s�raya g�re yap
                    advancedDataGridView1.Sort(advancedDataGridView1.Columns["T�KET�LECEK_�R�N_KODU"], ListSortDirection.Ascending);
                }
                else
                {
                    // S�ralamay� azalan s�raya g�re yap
                    advancedDataGridView1.Sort(advancedDataGridView1.Columns["T�KET�LECEK_�R�N_KODU"], ListSortDirection.Descending);
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            List<byte> durumu = new List<byte>() { 0, 1, 2 };
            comboBox1.DataSource = durumu;


            toolStripLabel1.ToolTipText =

                "1.�� Emri Kodu Gir\n" +

                "2.Tarih Se�\n" +

                "3.�� Emri Durumunu Se�\n" +

                "4.Verileri Listele\n" +

                "5.Evrak Tarihlerini Se�\n" +

                "6.Teslim Tarihlerini Se�\n" +

                "7.Talep olu�tur";


        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                // Ana form kapat�ld���nda t�m uygulamay� kapat
                Application.Exit();
            }
        }


        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            // DataGridView'deki verileri bir DataTable'a kopyalay�n
            DataTable dt = new DataTable();

            foreach (DataGridViewColumn column in advancedDataGridView1.Columns)
            {
                // E�er ValueType null ise, varsay�lan bir veri t�r� kullanabilirsiniz.
                Type columnType = column.ValueType ?? typeof(string);
                dt.Columns.Add(column.HeaderText, columnType);
            }

            // Sat�rlar� ekle
            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                DataRow dataRow = dt.NewRow();
                foreach (DataGridViewCell cell in row.Cells)
                {
                    dataRow[cell.ColumnIndex] = cell.Value;
                }
                dt.Rows.Add(dataRow);
            }

            // Excel uygulamas�n� ba�lat�n
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;

            // Yeni bir Excel �al��ma kitab� olu�turun
            Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];

            // DataTable'� Excel �al��ma sayfas�na aktar�n (tablo ba�l�klar�n� da dahil etmek i�in)
            int rowIndex = 1;

            // Ba�l�klar� yaz
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

            // Be�genin k��e noktalar�n� belirleyin
            Point[] pentagonPoints = new Point[]
            {
        new Point(button.Width / 2, 0),
        new Point(button.Width, button.Height / 3),
        new Point(button.Width * 2 / 3, button.Height),
        new Point(button.Width / 3, button.Height),
        new Point(0, button.Height / 3)
            };

            // Be�geni path'e ekleyin
            path.AddPolygon(pentagonPoints);

            // Path'i kullanarak yeni bir b�lge olu�turun
            button.Region = new Region(path);
        }

        private void fASON�MALATTALEB�ToolStripMenuItem_Click(object sender, EventArgs e)
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

        private void eSK�KODToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TALEB_MAL�YET TM = new TALEB_MAL�YET();
            TM.Show();
        }

        private void yEN�KODToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TALEP_MAL�YET_YEN�KOD_ TMYK = new TALEP_MAL�YET_YEN�KOD_();
            TMYK.Show();
        }

        private void bOMDAROTAD�ZELTMEToolStripMenuItem_Click(object sender, EventArgs e)
        {
            BOM_ROTA BR = new BOM_ROTA();
            BR.Show();
        }

        private void vER�LENTEKL�FMAL�YETHESABIToolStripMenuItem_Click(object sender, EventArgs e)
        {
            VER�LEN_TEKL�F_MAL�YET_HESABI VMH = new VER�LEN_TEKL�F_MAL�YET_HESABI();
            VMH.Show();
        }

        private void kODBULMAVEMAL�YETToolStripMenuItem_Click(object sender, EventArgs e)
        {
            KOD_BULMA KD = new KOD_BULMA();
            KD.Show();
        }

        private void s�RETRANSFER�ToolStripMenuItem_Click(object sender, EventArgs e)
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
                file.Filter = "Excel Dosyas� |*.xlsx";
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

                    // Checkbox s�tununu eklemiyoruz

                    // Verileri bir listeye aktar

                    foreach (DataRow row in data.Rows)
                    {
                        Dictionary<string, object> item = new Dictionary<string, object>();
                        foreach (DataColumn column in data.Columns)
                        {
                            item[column.ColumnName] = row[column]; // S�tun ad�n� anahtar, h�cre de�erini de�er olarak ekle
                        }
                        customerList.Add(item);
                    }

                    MessageBox.Show("BOM Listesi Y�klendi");
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

        private void tEKL�FVERMEToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TEKL�F TF = new TEKL�F();
            TF.Show();
        }

        private void lEADT�METoolStripMenuItem_Click(object sender, EventArgs e)
        {
            A�AC_TERM�N AT = new A�AC_TERM�N();
            AT.Show();
        }

        private void tALEPF�YATL�STES�ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TALEP_F�YAT TF = new TALEP_F�YAT();
            TF.Show();
        }

        private void kALANS�PAR��DETAYIToolStripMenuItem_Click(object sender, EventArgs e)
        {
            KALAN_S�PAR�S KS = new KALAN_S�PAR�S();
            KS.Show();
        }

        private void rOTAVE��MERKEZ��ZLEMEToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ROTA_�SMERKEZ� RI = new ROTA_�SMERKEZ�();
            RI.Show();
        }

        private void toolStripTextBox1_Click(object sender, EventArgs e)
        {
            Y�NET�M_EKRANI YE = new Y�NET�M_EKRANI();
            YE.Show();
        }

        private void mRP�NCES�TAHM�N�S�REToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MRP_ONCES�_TAHM�N�_SURE MRPOTS = new MRP_ONCES�_TAHM�N�_SURE();
            MRPOTS.Show();
        }
    }
}