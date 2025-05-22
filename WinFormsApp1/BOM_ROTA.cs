using Microsoft.Office.Interop.Excel;
using System.Data;
using System.Data.OleDb;
using System.Text.RegularExpressions;
using WinFormsApp1.Context;
using WinFormsApp1.Tables;
using ComboBox = System.Windows.Forms.ComboBox;
using Point = System.Drawing.Point;
using TextBox = System.Windows.Forms.TextBox;
namespace WinFormsApp1
{
    public partial class BOM_ROTA : Form
    {
        private MyDbContext dbContext;
        private SRFDbContext dbContextSRF;
        public BOM_ROTA()
        {
            dbContext = new MyDbContext();
            dbContextSRF = new SRFDbContext();
            InitializeComponent();
        }
        System.Data.DataTable data = new System.Data.DataTable();

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            // DataGridView'deki verileri bir DataTable'a kopyalayın
            System.Data.DataTable dt = new System.Data.DataTable();

            foreach (DataGridViewColumn column in advancedDataGridView1.Columns)
            {
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
                        dataRow[cell.ColumnIndex] = (cell.Value != null && (bool)cell.Value) ? "True" : "False";
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
                    worksheet.Cells[rowIndex, j + 1].NumberFormat = "@";
                    worksheet.Cells[rowIndex, j + 1] = dt.Rows[i][j].ToString();
                }
            }

            // "Malzeme Tanım" sütununu bul ve değerlerini Excel'in son sütununa yaz
            string malzemeTanımColumnName = "Malzeme Tanım"; // Sütun adı
            int malzemeTanımColumnIndex = -1;

            // "Malzeme Tanım" sütununu bul
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                if (dt.Columns[i].ColumnName == malzemeTanımColumnName)
                {
                    malzemeTanımColumnIndex = i;
                    break;
                }
            }

            // Eğer "Malzeme Tanım" sütunu bulunduysa, değerlerini yaz
            if (malzemeTanımColumnIndex != -1)
            {
                // H sütununu kaldır (8. sütun)
                worksheet.Columns[8].Delete(); // H sütunu 8. sütundur

                // Son sütuna başlık ekle
                worksheet.Cells[1, dt.Columns.Count] = malzemeTanımColumnName; // Başlık
                worksheet.Cells[1, dt.Columns.Count].Font.Bold = true; // Başlığı kalın yap

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    var malzemeTanımValue = dt.Rows[i][malzemeTanımColumnIndex].ToString();
                    worksheet.Cells[i + 2, dt.Columns.Count] = malzemeTanımValue; // Son sütuna yaz
                }
            }


        }


        private void UpdateDataTable()
        {
            // Check if "New QTY" column already exists
            if (!data.Columns.Contains("New QTY"))
            {
                // Assuming "New QTY" is of type string, change it according to your actual data type
                data.Columns.Add("New QTY", typeof(double));
            }

            // Check if "Aciklama" column already exists
            if (!data.Columns.Contains("Aciklama"))
            {
                // Assuming "Aciklama" is of type string, change it according to your actual data type
                data.Columns.Add("Aciklama", typeof(string));
            }

            // Get distinct Module_No values
            var distinctModuleNos = data.AsEnumerable()
                .Select(row => row.Field<string>("Module No"))
                .Where(moduleNo => !string.IsNullOrEmpty(moduleNo))
                .Distinct()
                .ToList();

            foreach (var moduleNo in distinctModuleNos)
            {
                var moduleRows = data.AsEnumerable()
                    .Where(row => row.Field<string>("Module No") == moduleNo)
                    .ToList();

                string previousPartNumber = null;

                for (int i = 0; i < moduleRows.Count; i++)
                {
                    string currentItem = moduleRows[i].Field<string>("Item");

                    if (currentItem != null)
                    {
                        string[] itemParts = currentItem.Split('.');
                        string itemQtyString = moduleRows[i].Field<string>("Item QTY");

                        double currentQty;
                        // Try to parse string to double
                        if (!double.TryParse(itemQtyString, out currentQty))
                        {
                            // Parsing failed, handle the error or set a default value
                            currentQty = 0; // Or any default value you choose
                        }

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
                                    double parentItemQty;
                                    string parentItemQtyString = moduleRows[parentRowIndex].Field<string>("Item QTY");

                                    // Try to parse string to double
                                    if (!double.TryParse(parentItemQtyString, out parentItemQty))
                                    {
                                        // Parsing failed, handle the error or set a default value
                                        parentItemQty = 0; // Or any default value you choose
                                    }

                                    parentQty *= parentItemQty;

                                    // Check for null values before calculations
                                    if (!double.IsNaN(parentItemQty))
                                    {
                                        previousPartNumber = moduleRows[parentRowIndex].Field<string>("Part Number");
                                    }
                                }
                            }

                            // Update the "New QTY" column for the current row
                            moduleRows[i].SetField("New QTY", parentQty);

                            // Update the "Aciklama" column
                            moduleRows[i].SetField("Aciklama", previousPartNumber);
                        }
                    }
                }
            }

            // AdvancedDataGridView'e ata
            advancedDataGridView1.DataSource = data;
        }

        private void button1_Click(object sender, EventArgs e)
        {

            try
            {
                OpenFileDialog file = new OpenFileDialog();
                file.Filter = "Excel Dosyası |*.xlsx";
                file.ShowDialog();

                string tamYol = file.FileName;

                string baglantiAdresi = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + tamYol + ";Extended Properties='Excel 12.0;IMEX=1;'";

                OleDbConnection baglanti = new OleDbConnection(baglantiAdresi);

                OleDbCommand komut = new OleDbCommand("Select * From [" + textBox2.Text + "$]", baglanti);

                baglanti.Open();

                OleDbDataAdapter da = new OleDbDataAdapter(komut);

                da.Fill(data);

                // "Ham Agirlik" ve "Mass" sütunlarını kontrol et ve virgülü nokta ile değiştir
                foreach (DataRow row in data.Rows)
                {
                    // "Ham Agirlik" sütunu için kontrol
                    if (row["Ham Agirlik"] != DBNull.Value)
                    {
                        string hamAgirlikStr = row["Ham Agirlik"].ToString();
                        if (hamAgirlikStr.Contains(","))
                        {
                            // Virgül varsa, onu noktaya çevir
                            hamAgirlikStr = hamAgirlikStr.Replace(",", ".");
                        }
                        row["Ham Agirlik"] = hamAgirlikStr; // Değeri güncelle
                    }

                    // "Mass" sütunu için kontrol
                    if (row["Mass"] != DBNull.Value)
                    {
                        string massStr = row["Mass"].ToString();
                        if (massStr.Contains(","))
                        {
                            // Virgül varsa, onu noktaya çevir
                            massStr = massStr.Replace(",", ".");
                        }
                        row["Mass"] = massStr; // Değeri güncelle
                    }
                }

                // Veri kaynağını DataGridView'e ata
                advancedDataGridView1.DataSource = data;

                MessageBox.Show("Bom Listesi Yüklendi");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }




        }



        private void advancedDataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0 && e.RowIndex >= 0) // İlk sütun ve satırın başlığı değilse
            {
                DataGridViewCheckBoxCell checkbox = (DataGridViewCheckBoxCell)advancedDataGridView1.Rows[e.RowIndex].Cells[0]; // İlk sütun
                checkbox.Value = !(bool)checkbox.Value; // Checkbox durumunu tersine çevir

                // DataGridView'daki işaretlemenin DataTable'a yansıtılması
                DataRowView rowView = (DataRowView)advancedDataGridView1.Rows[e.RowIndex].DataBoundItem;
                DataRow row = rowView.Row;
                row[0] = checkbox.Value; // İlk sütun

                // DataGridView'i güncelle
                advancedDataGridView1.Refresh(); // Yenileme işlemi
            }
        }

        private void checkBox1_Click(object sender, EventArgs e)
        {

            // DataGridView'de dolaşarak tüm hücreleri true yap
            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                DataGridViewCheckBoxCell checkBoxCell = row.Cells["checkboxColumn"] as DataGridViewCheckBoxCell;
                checkBoxCell.Value = checkBoxCell.Value == null || !(bool)checkBoxCell.Value;
            }
        }

        private void BOM_ROTA_Load(object sender, EventArgs e)
        {
            toolStripLabel1.ToolTipText =

                "1.Excel Sayfa İsmini Gir\n" +

                "2.Listele\n" +

                "3.Part Number sütununda Filtreleme Yap\n" +

                "4.Hepsini İşaretleyi Aktif Et\n" +

                "5.Güncelleme Yap\n" +

                "6.Filtrelemeyi Kaldır\n" +

                "7.Sil Yap\n" +

                "8.CheckboxColumn sütununda 'TRUE' ları filtrele\n" +

                "9.İşareti Kaldır\n" +

                "10.Excele Aktar";
        }
        private void button5_Click(object sender, EventArgs e)
        {
            var sarfKodList = dbContextSRF.SARF_MALZEME_KOD.Select(x => x.KODU).ToList();

            for (int i = advancedDataGridView1.Rows.Count - 1; i >= 0; i--)
            {
                DataGridViewRow row = advancedDataGridView1.Rows[i];
                if (row.IsNewRow) continue;

                string partNumber = row.Cells["Part Number"].Value?.ToString();

                if (sarfKodList.Contains(partNumber))
                {
                    advancedDataGridView1.Rows.RemoveAt(i);
                }
            }

            MessageBox.Show("Eşleşen satırlar silindi!");

        }

        private void button6_Click(object sender, EventArgs e)
        {
            // Kaldırılacak satırları tutmak için liste
            var rowsToRemove = new List<DataGridViewRow>();

            // Her satırı döngüye al
            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                string kaynakValue = row.Cells["Fason"].Value?.ToString();
                string? materialValue = row.Cells["Material"].Value?.ToString();

                if (kaynakValue == "EVET" && (materialValue == "KAYNAK" || string.IsNullOrEmpty(materialValue)))
                {
                    string kaynakItem = row.Cells["Item"].Value?.ToString();
                    if (!string.IsNullOrEmpty(kaynakItem))
                    {
                        // Eşleşen satırları bul
                        var matchingRows = advancedDataGridView1.Rows
                            .Cast<DataGridViewRow>()
                            .Where(r =>
                            {
                                string itemValue = r.Cells["Item"].Value?.ToString();
                                if (string.IsNullOrEmpty(itemValue))
                                    return false;

                                // Kaynak item'ın bir alt seviyesi olup olmadığını kontrol et
                                string pattern = $"^{Regex.Escape(kaynakItem)}\\.\\d+$";
                                return Regex.IsMatch(itemValue, pattern);
                            }).ToList();

                        rowsToRemove.AddRange(matchingRows);
                    }
                }
            }

            // Satırları kaldır
            foreach (DataGridViewRow rowToRemove in rowsToRemove)
            {
                if (advancedDataGridView1.Rows.Contains(rowToRemove))
                {
                    advancedDataGridView1.Rows.Remove(rowToRemove);
                }
            }
        }



        private void button2_Click_1(object sender, EventArgs e)
        {

            // DataTable oluştur
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("Ürün Kodu");
            dt.Columns.Add("İş Emri Kodu");
            dt.Columns.Add("Durum");
            dt.Columns.Add("Proje Kodu");
            dt.Columns.Add("Oluşma Tarihi");
            dt.Columns.Add("bağlı olduğu");
            // Daha sade tarih tanımı
            var baslangic = DateTime.Parse("2024-01-01");
            var bitis = DateTime.Parse("2024-12-31 23:59:59");

            // Verileri çek
            var uplUrStokKodAnd0 = dbContext.URETIM_MALZEME_PLANLAMA
                .Join(dbContext.ISEMIRLERI,
                    uretim => uretim.upl_isemri,
                    isEmir => isEmir.is_Kod,
                    (uretim, isEmir) => new { Uretim = uretim, IsEmir = isEmir })
                .Where(joined => joined.IsEmir.is_EmriDurumu == 1&&joined.Uretim.upl_uretim_tuket==1
                    && joined.Uretim.upl_kodu.EndsWith(".128") && joined.IsEmir.is_create_date >= baslangic
        && joined.IsEmir.is_create_date <= bitis)
                .Select(joined => new
                {
                    joined.Uretim.upl_kodu,
                    joined.IsEmir.is_Kod,
                    joined.IsEmir.is_EmriDurumu,
                    joined.IsEmir.is_ProjeKodu,
                    joined.IsEmir.is_create_date,
                    joined.IsEmir.is_BagliOlduguIsemri

                })
                .ToList();

            // Listeyi doldur
            foreach (var item in uplUrStokKodAnd0)
            {
                dt.Rows.Add(item.upl_kodu, item.is_Kod, item.is_EmriDurumu, item.is_ProjeKodu, item.is_create_date.Date,item.is_BagliOlduguIsemri);
            }

            // AdvancedDataGridView'e ata
            advancedDataGridView1.DataSource = dt;


        }

        private void button3_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                if (row.IsNewRow) continue;

                string isKod = row.Cells["İş Emri Kodu"].Value?.ToString();
                string projeKod = row.Cells["Proje Kodu"].Value?.ToString();
                string bagliOldugu = row.Cells["bağlı olduğu"].Value?.ToString();

                if (!string.IsNullOrEmpty(isKod))
                {
                    // Eşleşen kaydı bul
                    var isEmri = dbContext.ISEMIRLERI
                        .FirstOrDefault(x =>
                            x.is_Kod == isKod &&
                            x.is_ProjeKodu == projeKod);

                    // Varsa güncelle
                    if (isEmri != null)
                    {
                        isEmri.is_EmriDurumu = 2;
                        dbContext.Entry(isEmri).Property(x => x.is_EmriDurumu).IsModified = true;
                    }
                }
                if (!string.IsNullOrEmpty(isKod))
                {
                    // Eşleşen kaydı bul
                    var isEmri = dbContext.ISEMIRLERI
                        .FirstOrDefault(x =>
                            x.is_Kod == bagliOldugu &&
                            x.is_ProjeKodu == projeKod);

                    // Varsa güncelle
                    if (isEmri != null)
                    {
                        isEmri.is_EmriDurumu = 2;
                        dbContext.Entry(isEmri).Property(x => x.is_EmriDurumu).IsModified = true;
                    }
                }
            }

            // Değişiklikleri kaydet
            dbContext.SaveChanges();

            MessageBox.Show("Güncelleme Başarılı...");

        }
    }
}
