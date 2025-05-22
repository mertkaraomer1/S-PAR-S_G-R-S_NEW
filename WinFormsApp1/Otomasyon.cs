using Microsoft.Data.SqlClient;
using System.Data;
using System.Data.OleDb;
using WinFormsApp1.Tables;
using DataTable = System.Data.DataTable;

namespace WinFormsApp1
{
    public partial class Otomasyon : Form
    {
        public Otomasyon()
        {
            InitializeComponent();
        }
        List<Malzeme> malzemeList = new List<Malzeme>();
        List<VeriCek> dataList = new List<VeriCek>();
        private void Otomasyon_Load(object sender, EventArgs e)
        {
            toolStripLabel1.ToolTipText =

                "1.Excel Sayfa İsmini Gir\n" +

                "2.Excel Aktar\n" +

                "3.Veritabanından Veri Çek\n" +

                "4.Listele";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //dataGridView1.Rows.Clear();
            try
            {
                // Dosya seçme penceresi açmak için
                OpenFileDialog file = new OpenFileDialog();
                file.Filter = "Excel Dosyası |*.xlsx";
                file.ShowDialog();

                // seçtiğimiz excel'in tam yolu
                string tamYol = file.FileName;

                //Excel bağlantı adresi
                string baglantiAdresi = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + tamYol + ";Extended Properties='Excel 12.0;IMEX=1;'";

                //bağlantı 
                OleDbConnection baglanti = new(baglantiAdresi);

                //tüm verileri seçmek için select sorgumuz. Sayfa1 olan kısmı Excel'de hangi sayfayı açmak istiyosanız orayı yazabilirsiniz.
                OleDbCommand komut = new OleDbCommand("Select * From [" + textBox1.Text + "$]", baglanti);

                //bağlantıyı açıyoruz.
                baglanti.Open();

                //Gelen verileri DataAdapter'e atıyoruz.
                OleDbDataAdapter da = new OleDbDataAdapter(komut);

                //Grid'imiz için bir DataTable oluşturuyoruz.
                System.Data.DataTable data = new System.Data.DataTable();

                //DataAdapter'da ki verileri data adındaki DataTable'a dolduruyoruz.
                da.Fill(data);

                //DataGrid'imizin kaynağını oluşturduğumuz DataTable ile dolduruyoruz.
                //dataGridView1.DataSource = data;

                malzemeList = DataTableConverter.ConvertTo<Malzeme>(data);
                MessageBox.Show("Malzeme Listesi Yüklendi");
            }
            catch (Exception ex)
            {
                // Hata alırsak ekrana bastırıyoruz.
                MessageBox.Show(ex.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // SqlConnection oluştur ve aç
            string connectionString = "Data Source=SRVMIKRO;Initial Catalog=MikroDB_V16_ICM;Integrated Security=True;Connect Timeout=10;Encrypt=False;Trust Server Certificate=False;Application Intent=ReadWrite;Multi Subnet Failover=False";


            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                string sqlQuery = @"SELECT
                            sto_kod AS [stok kodu] /* KODU */,
                            sto_isim AS [stok adı] /* ADI */,
                            sto_marka_kodu AS 'MARKA KODU',
                            sto_birim1_ad AS 'BİRİM1',
                            dbo.fn_DepodakiMiktar(sto_kod, 73151, '') AS 'ADET'
                        FROM dbo.STOKLAR WITH (NOLOCK)
                        LEFT OUTER JOIN dbo.STOK_SATIS_FIYATLARI_F1_D0_VIEW ON sfiyat_stokkod = sto_kod AND sfiyat_satirno = 1
                        LEFT OUTER JOIN MikroDB_V16.dbo.KUR_ISIMLERI ON Kur_No = ISNULL(sfiyat_doviz, 0)
                        LEFT OUTER JOIN STOKLAR_USER ON STOKLAR_USER.Record_uid = sto_Guid
                        ORDER BY sto_kod";

                using (SqlCommand command = new SqlCommand(sqlQuery, connection))
                {
                    using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                    {
                        DataTable dataTable = new DataTable();
                        adapter.Fill(dataTable);

                        // DataTable'daki verileri liste içine aktar
                        foreach (DataRow row in dataTable.Rows)
                        {
                            VeriCek item = new VeriCek
                            {
                                StokKodu = row["stok kodu"].ToString(),
                                StokAdi = row["stok adı"].ToString(),
                                MarkaKodu = row["MARKA KODU"].ToString(),
                                Birim1 = row["BİRİM1"].ToString(),
                                DepoADET = row["ADET"].ToString()
                            };
                            dataList.Add(item);
                        }

                    }
                    MessageBox.Show("Depo Malzeme Listesi Yüklendi");
                }
            }

            // Şimdi dataList adlı listeniz verileri içeriyor.
            // Bu listeden verilere erişebilirsiniz.
        }
        int sayac = 0;
        private void button3_Click(object sender, EventArgs e)
        {// Eşleşenleri bulun
            var eslesmisVeriler = from malzeme in malzemeList
                                  join data in dataList on malzeme.StokKodu equals data.StokKodu into joinedData
                                  from data in joinedData.DefaultIfEmpty()
                                  group new { malzeme, data } by malzeme.StokKodu into groupedData
                                  select new
                                  {
                                      StokKodu = groupedData.Key,
                                      İhtiyacMiktar = groupedData.Sum(x => Convert.ToInt32(x.malzeme.İhtiyacMiktar)),
                                      MalAdi = groupedData.First().malzeme.MalAdi,
                                      DepoADET = Convert.ToInt32(groupedData.First().data != null ? groupedData.First().data.DepoADET : 0)
                                  };


            // Eşleşmeyenleri bulun
            var eslesmeyenVeriler = from malzeme in malzemeList
                                    where !eslesmisVeriler.Any(x => x.StokKodu == malzeme.StokKodu)
                                    select new
                                    {
                                        StokKodu = malzeme.StokKodu,
                                        İhtiyacMiktar = Convert.ToInt32(malzeme.İhtiyacMiktar),
                                        MalAdi = malzeme.MalAdi,
                                        DepoADET = 0 // veya istediğiniz bir değeri buraya ekleyebilirsiniz
                                    };

            // Eşleşen ve eşleşmeyen verileri birleştirin
            var combinedDataList = eslesmisVeriler.Concat(eslesmeyenVeriler).ToList();

            // DataGridView'e verileri yükleyin
            dataGridView1.DataSource = combinedDataList;

            // İlk sütunda satır numarasını göstermek için DataGridView'e bir sütun ekleyin
            dataGridView1.Columns.Insert(0, new DataGridViewTextBoxColumn()
            {
                HeaderText = "Sıra No",
                Name = "RowNumber"
            });

            // Sipariş verilecek miktar sütununu ekleyin
            dataGridView1.Columns.Add(new DataGridViewTextBoxColumn()
            {
                HeaderText = "Verilecek Sipariş Miktarı",
                Name = "SiparisMiktari"
            });

            // Her satırın başındaki hücrelere sıra numarasını ekleyin ve sipariş miktarını ayarlayın
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                dataGridView1.Rows[i].Cells["RowNumber"].Value = (i + 1).ToString();
                var adetMiktari = Convert.ToInt32(dataGridView1.Rows[i].Cells["DepoADET"].Value);
                var miktar = Convert.ToInt32(dataGridView1.Rows[i].Cells["İhtiyacMiktar"].Value);

                if (miktar > adetMiktari)
                {
                    dataGridView1.Rows[i].Cells["SiparisMiktari"].Value = miktar - adetMiktari;
                }
                else
                {
                    dataGridView1.Rows[i].Cells["SiparisMiktari"].Value = adetMiktari >= miktar ? 0 : miktar;
                }
            }

        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            // DataGridView'deki verileri bir DataTable'a kopyalayın
            DataTable dt = new DataTable();

            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                // Eğer ValueType null ise, varsayılan bir veri türü kullanabilirsiniz.
                Type columnType = column.ValueType ?? typeof(string);
                dt.Columns.Add(column.HeaderText, columnType);
            }

            // Satırları ekle
            foreach (DataGridViewRow row in dataGridView1.Rows)
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
    }
}
