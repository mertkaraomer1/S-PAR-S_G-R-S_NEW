using System.Data;
using System.Data.OleDb;
using System.Drawing.Drawing2D;
using WinFormsApp1.Context;
using WinFormsApp1.Tables;
using ComboBox = System.Windows.Forms.ComboBox;
using TextBox = System.Windows.Forms.TextBox;

namespace WinFormsApp1
{
    public partial class EXCEL_TALEP : Form
    {
        private MyDbContext dbContext;
        public EXCEL_TALEP()
        {
            dbContext = new MyDbContext();
            InitializeComponent();

            DataGridViewCheckBoxColumn checkBoxColumn1 = new DataGridViewCheckBoxColumn();
            checkBoxColumn1.HeaderText = "T";
            checkBoxColumn1.Name = "checkBoxColumn1";
            advancedDataGridView1.Columns.Insert(0, checkBoxColumn1);
        }
        DataTable table = new DataTable();
        private void UpdateDataTable()
        {
            // Check if "New QTY" column already exists
            if (!table.Columns.Contains("New QTY"))
            {
                // Assuming "New QTY" is of type string, change it according to your actual data type
                table.Columns.Add("New QTY", typeof(double));
            }

            // Check if "Aciklama" column already exists
            if (!table.Columns.Contains("Aciklama"))
            {
                // Assuming "Aciklama" is of type string, change it according to your actual data type
                table.Columns.Add("Aciklama", typeof(string));
            }

            // Get distinct Module_No values
            var distinctModuleNos = table.AsEnumerable()
                .Select(row => row.Field<string>("Module No"))
                .Where(moduleNo => !string.IsNullOrEmpty(moduleNo))
                .Distinct()
                .ToList();

            foreach (var moduleNo in distinctModuleNos)
            {
                var moduleRows = table.AsEnumerable()
                    .Where(row => row.Field<string>("Module No") == moduleNo)
                    .ToList();

                string previousPartNumber = null;

                for (int i = 0; i < moduleRows.Count; i++)
                {
                    string currentItem = moduleRows[i].Field<string>("Item");

                    if (currentItem != null)
                    {
                        string[] itemParts = currentItem.Split('.');
                        double currentQty = moduleRows[i].Field<double>("Item QTY");

                        // Check for null values before calculations
                        if (!double.IsNaN(currentQty))
                        {
                            // Çarpmaya başlamadan önce her seferinde parentQty'yi sıfırla
                            double parentQty = 1;

                            for (int j = 0; j < itemParts.Length; j++)
                            {
                                // Find the parent item's row index
                                string parentItem = string.Join(".", itemParts.Take(j + 1));
                                int parentRowIndex = moduleRows.FindLastIndex(r => r.Field<string>("Item") == parentItem);

                                if (parentRowIndex != -1)
                                {
                                    double parentItemQty = moduleRows[parentRowIndex].Field<double>("Item QTY");
                                    parentQty *= parentItemQty;

                                    // Check for null values before calculations
                                    if (!double.IsNaN(parentItemQty))
                                    {
                                        previousPartNumber = moduleRows[parentRowIndex].Field<string>("Part Number");
                                    }
                                }
                            }

                            // Update the "New QTY" column for the current row
                            moduleRows[i].SetField("New QTY", parentQty.ToString());

                            // Update the "Aciklama" column
                            moduleRows[i].SetField("Aciklama", previousPartNumber);
                        }
                    }
                }
            }

            // AdvancedDataGridView'e ata
            advancedDataGridView1.DataSource = table;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                // Dosya seçme penceresi açmak için
                OpenFileDialog file = new OpenFileDialog();
                file.Filter = "Excel Dosyası |*.xlsx";
                file.ShowDialog();

                // seçtiğimiz excel'in tam yolu
                string tamYol = file.FileName;

                // Excel bağlantı adresi
                string baglantiAdresi = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + tamYol + ";Extended Properties='Excel 12.0;IMEX=1;'";

                // Bağlantı oluştur
                using (OleDbConnection baglanti = new OleDbConnection(baglantiAdresi))
                {
                    // Tüm verileri seçmek için select sorgumuz. Sayfa1 olan kısmı Excel'de hangi sayfayı açmak istiyorsanız orayı yazabilirsiniz.
                    using (OleDbCommand komut = new OleDbCommand("Select * From [" + textBox2.Text + "$]", baglanti))
                    {
                        // Bağlantıyı açıyoruz.
                        baglanti.Open();

                        // Gelen verileri DataAdapter'e atıyoruz.
                        using (OleDbDataAdapter da = new OleDbDataAdapter(komut))
                        {
                            // DataAdapter'da ki verileri data adındaki DataTable'a dolduruyoruz.
                            da.Fill(table);
                        }
                    }
                }

                // İşlemleri gerçekleştir ve AdvancedDataGridView'e ata
                UpdateDataTable();

                MessageBox.Show("Bom Listesi Yüklendi");
            }
            catch (Exception ex)
            {
                // Hata alırsak ekrana bastırıyoruz.
                MessageBox.Show(ex.Message);
            }

        }
        private void button1_Click(object sender, EventArgs e)
        {
            // Get distinct Module_No values
            var distinctModuleNos = advancedDataGridView1.Rows
                .Cast<DataGridViewRow>()
                .Select(row => row.Cells["Module No"].Value?.ToString())
                .Where(moduleNo => !string.IsNullOrEmpty(moduleNo))
                .Distinct()
                .ToList();

            foreach (var moduleNo in distinctModuleNos)
            {
                // Filter rows based on the current Module_No
                var moduleRows = advancedDataGridView1.Rows
                    .Cast<DataGridViewRow>()
                    .Where(row => row.Cells["Module No"].Value?.ToString() == moduleNo)
                    .ToList();

                for (int i = 0; i < moduleRows.Count; i++)
                {
                    string currentItem = moduleRows[i].Cells["Item"].Value?.ToString();

                    if (currentItem != null && currentItem.Contains("."))
                    {
                        // If it's a sub-item (contains a dot), find its parent item
                        string parentItem = currentItem.Substring(0, currentItem.LastIndexOf('.'));

                        // Find the parent item's row index
                        int parentRowIndex = moduleRows.FindLastIndex(row => row.Cells["Item"].Value?.ToString() == parentItem);

                        if (parentRowIndex != -1)
                        {
                            // Find the parent item's quantity and part number
                            int parentQty = int.TryParse(moduleRows[parentRowIndex].Cells["New QTY"].Value?.ToString(), out var qty) ? qty : 0;
                            string parentPartNumber = moduleRows[parentRowIndex].Cells[2].Value?.ToString() ?? ""; // Assuming Part Number is in the 2nd column

                            // Find the current item's quantity
                            int currentQty = int.TryParse(moduleRows[i].Cells[3].Value?.ToString(), out var itemQty) ? itemQty : 0; // Replace with the actual index of "Item QTY" column

                            // Calculate the new quantity
                            int newQty = parentQty * currentQty;

                            // Update the "New QTY" column
                            moduleRows[i].Cells["New QTY"].Value = newQty.ToString();

                            // Update the "New Part Number" column
                            moduleRows[i].Cells["Aciklama"].Value = parentPartNumber;
                        }
                    }
                    else
                    {
                        // If it's not a sub-item, copy the "Item QTY" to "New QTY" and copy Part Number to "New Part Number"
                        moduleRows[i].Cells["New QTY"].Value = moduleRows[i].Cells[3].Value?.ToString() ?? "";
                        moduleRows[i].Cells["Aciklama"].Value = moduleRows[i].Cells[4].Value?.ToString() ?? ""; // Assuming Part Number is in the 2nd column
                    }
                }
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            EkleTalepler();
            IsemriGüncelle();


            // Mesaj kutularını göster
            MessageBox.Show("Veriler başarıyla eklendi ve güncellendi.");
        }
        public void EkleTalepler()
        {
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
                "34",
                "25",
                "28"
            });

            TextBox textBox = new TextBox
            {
                Location = new Point(20, 60),
                Width = 200,
                Visible = false  // Başlangıçta görünmez
            };
            TextBox textBoxBelge = new TextBox
            {
                Location = new Point(50, 100),
                Width = 200,
                Visible = false  // Başlangıçta görünmez
            };
            Button onayButton = new Button
            {
                Text = "Onay",
                Location = new Point(100, 100),
                Size = new Size(60, 30),
            };
            // ComboBox'ın SelectedIndexChanged etkinliği
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
                    textBoxBelge.Text = "SŞ";
                }
                else if (selectedIndex == 2)
                {
                    textBox.Text = "01-09-02-03";
                    textBoxBelge.Text = "MK";
                }
                else if (selectedIndex == 3)
                {
                    textBox.Text = "01-09-02-10";
                    textBoxBelge.Text = "BŞ";
                }
                else if (selectedIndex == 4)
                {
                    textBox.Text = "01-08-02-06";
                    textBoxBelge.Text = "TK";
                }
                else if (selectedIndex == 5)
                {
                    textBox.Text = "01-08-02-04";
                    textBoxBelge.Text = "ŞB";
                }
            };
            onayButton.Click += (s, e) =>
            {
                List<SATINALMA_TALEPLERI> talepListesi = new List<SATINALMA_TALEPLERI>();
                int sayac = 0;
                int nextStlEvrakSira = dbContext.SATINALMA_TALEPLERI.Max(item => item.stl_evrak_sira);
                int evrakSira = nextStlEvrakSira + 1; // Benzersiz bir şekilde ayarlanmalıdır
                int satirNo = sayac;                         // DataGridView'deki her bir satır için
                var selectedDate1 = dateTimePicker2.Value.Date;
                var selectedDate = dateTimePicker1.Value.Date;
                foreach (DataGridViewRow row in advancedDataGridView1.Rows)
                {
                    if (!row.IsNewRow)
                    {
                        bool isSelected = Convert.ToBoolean(row.Cells["checkBoxColumn1"].Value);
                        string Item = row.Cells["Item"].Value.ToString();
                        string Part_Number = row.Cells["Part Number"].Value.ToString();
                        int Item_QTY = Convert.ToInt32(row.Cells["Item QTY"].Value.ToString());
                        string Module_No = row.Cells["Module No"].Value.ToString();
                        int New_QTY = Convert.ToInt32(row.Cells["New QTY"].Value.ToString());
                        string Aciklama = row.Cells["Aciklama"].Value.ToString();
                        if (isSelected)
                        {
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
                                stl_special1 = "35",
                                stl_special2 = "",
                                stl_special3 = "",
                                stl_firmano = 0,
                                stl_subeno = 0,
                                stl_tarihi = DateTime.Now.Date,
                                stl_teslim_tarihi = DateTime.Now.Date,
                                stl_evrak_seri = "T",
                                stl_evrak_sira = evrakSira,
                                stl_satir_no = sayac,
                                stl_belge_no = textBoxBelge.Text,
                                stl_belge_tarihi = DateTime.Now.Date,
                                stl_Sor_Merk = textBox1.Text,
                                stl_Stok_kodu = Part_Number,
                                stl_Satici_Kodu = "",
                                stl_miktari = New_QTY,
                                stl_teslim_miktari = 0,
                                stl_aciklama = Aciklama + ".01.014",
                                stl_aciklama2 = "",
                                stl_depo_no = 70150,
                                stl_kapat_fl = 0,
                                stl_projekodu = textBox1.Text,
                                stl_parti_kodu = "",
                                stl_lot_no = "0",
                                stl_OnaylayanKulNo = "0",
                                stl_cagrilabilir_fl = 1,
                                stl_harekettipi = 0,
                                stl_talep_eden = textBox.Text,
                                stl_kapatmanedenkod = "",
                                stl_KaynakSip_uid = Guid.Empty,
                                stl_HareketGrupKodu1 = "",
                                stl_HareketGrupKodu2 = "",
                                stl_HareketGrupKodu3 = "",
                                stl_birim_pntr = 1,
                                stl_Franchise_Fiyati = 0,
                            };

                            talepListesi.Add(talep1);
                            sayac++;
                        }
                        if (isSelected)
                        {
                            var sorgu = dbContext.ISEMIRLERI
                                .Join(
                                    dbContext.URETIM_MALZEME_PLANLAMA,
                                    isEmir => isEmir.is_Kod,
                                    uretim => uretim.upl_isemri,
                                    (isEmir, uretim) => new { isEmir, uretim }
                                )
                                .Where(x => x.isEmir.is_ProjeKodu == textBox1.Text &&
                                            (x.isEmir.is_create_date.Date >= selectedDate && x.isEmir.is_create_date.Date <= selectedDate1))
                                .Select(x => new
                                {
                                    PROJE_KOD = x.isEmir.is_ProjeKodu,
                                    İŞ_EMRİ_KODU = x.isEmir.is_Kod,
                                    UPL_KODU = x.uretim.upl_kodu,
                                    İŞEMRİ_DURUMU = x.isEmir.is_EmriDurumu
                                })
                                .ToList();

                            foreach (var item in sorgu)
                            {
                                if (item.UPL_KODU == Part_Number)
                                {
                                    // Eşleşen Part_Number'ın iş emir durumunu güncelleme
                                    var isEmirToUpdate = dbContext.ISEMIRLERI.FirstOrDefault(i => i.is_Kod == item.İŞ_EMRİ_KODU);

                                    if (isEmirToUpdate != null)
                                    {
                                        isEmirToUpdate.is_EmriDurumu = 0;
                                    }
                                }
                            }

                        }


                    }
                }

                dbContext.SATINALMA_TALEPLERI.AddRange(talepListesi);
                dbContext.SaveChanges();
                sebepSecimiForm.Close();
            };
            // Forma kontrolleri ekle
            sebepSecimiForm.Controls.AddRange(new Control[] { comboBox, textBox, textBoxBelge, onayButton });

            // Formu görünür yap
            sebepSecimiForm.ShowDialog();
        }
        public void IsemriGüncelle()
        {
            var selectedDate1 = dateTimePicker2.Value.Date;
            var selectedDate = dateTimePicker1.Value.Date;
            string projeKodu = textBox1.Text;

            // ISEMIRLERI tablosu ile URETIM_MALZEME_PLANLAMA tablosunu join et
            var sorgu = from isEmir in dbContext.ISEMIRLERI
                        join uretim in dbContext.URETIM_MALZEME_PLANLAMA on isEmir.is_Kod equals uretim.upl_isemri
                        where isEmir.is_ProjeKodu == projeKodu &&
                              isEmir.is_create_date.Date >= selectedDate &&
                              isEmir.is_create_date.Date <= selectedDate1
                        select new
                        {
                            UplKodu = uretim.upl_kodu,
                            IsEmriDurumu = isEmir.is_EmriDurumu,
                            IsKod = isEmir.is_Kod
                        };

            // Sorgu sonucundaki verileri liste olarak al
            var veriListesi = sorgu.ToList();

            // Verileri kullanarak istediğiniz işlemleri gerçekleştirin
            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                if (!row.IsNewRow)
                {
                    bool isSelected = Convert.ToBoolean(row.Cells["checkBoxColumn1"].Value);

                    if (isSelected)
                    {
                        string PARTNUMBER = row.Cells["Part Number"].Value.ToString();

                        // sorgu verileri ile eşleşen ÜRETİLECEK_ÜRÜN_KODU'nu bul
                        var eşleşenVeri = veriListesi
                            .Where(veri =>
                                veri.UplKodu.Length >= 13 &&
                                veri.UplKodu.Substring(0, 13) == PARTNUMBER)
                            .Select(x => x.IsKod)
                            .ToList();

                        foreach (string isEmriKodu in eşleşenVeri)
                        {
                            // ISEMIRLERI tablosundan ilgili satırı bul
                            var isEmirToUpdate = dbContext.ISEMIRLERI.FirstOrDefault(isEmir => isEmir.is_Kod == isEmriKodu);

                            if (isEmirToUpdate != null)
                            {
                                // is_EmriDurumu'nu 0 olarak güncelle
                                isEmirToUpdate.is_EmriDurumu = 0;

                                // Değişiklikleri kaydet
                                dbContext.SaveChanges();
                            }
                        }
                    }
                }
            }
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
        private void checkBox1_Click(object sender, EventArgs e)
        {
            // DataGridView'de dolaşarak tüm hücreleri true yap
            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                DataGridViewCheckBoxCell checkBoxCell = row.Cells["checkBoxColumn1"] as DataGridViewCheckBoxCell;
                checkBoxCell.Value = true;
            }
        }

        private void EXCEL_TALEP_Load(object sender, EventArgs e)
        {
            toolStripLabel1.ToolTipText =
                "Excel formatı\n" +

         "1.sütun=Item\n" +

         "2.sütun=Part Number\n" +

         "3.sütun=Item QTY\n" +

         "4.sütun=Module No\n" +

         "Kullanım Şekli\n" +

         "1.Excel aktarıldıktan sonra Güncelle butonuna bas\n" +

         "2.Filtreleme yap\n" +

         "3.Hepsini işaretleyi aktif et\n" +

         "4.Tarih seç ve Proje kodu yaz\n" +

         "5.Talep oluştur";

        }


    }
}
