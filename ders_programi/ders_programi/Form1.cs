using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;



namespace ders_programi
{
    public partial class Form1 : Form
    {

        // database bağlantısı
        private string connectionString = "Data Source=.;Initial Catalog=ders_p;Integrated Security=True";

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
            this.dersTableAdapter1.Fill(this.ders_pDers2.ders);
            
            this.sinifTableAdapter.Fill(this.ders_pDataSet.sinif);
            
            this.ogretmenTableAdapter.Fill(this.ders_pDataOgr.ogretmen);
            
            this.dersTableAdapter.Fill(this.ders_pDataDers.ders);
            
            this.dataTable1TableAdapter.Fill(this.ders_pDataSet1.DataTable1);
            
            this.dataTable1TableAdapter.Fill(this.ders_pDataSet1.DataTable1);
            


        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (comboBox7.SelectedIndex != -1)
            {
                DataRowView selectedRow = (DataRowView)comboBox7.SelectedItem;
                int primaryKeyValue = (int)selectedRow["ders_id"];

                try
                {
                    using (SqlConnection connection = new SqlConnection(connectionString))
                    {
                        connection.Open();

                        using (SqlCommand command = new SqlCommand("DELETE FROM ders WHERE ders_id = @ders_id", connection))
                        {
                            command.Parameters.AddWithValue("@ders_id", primaryKeyValue);
                            command.ExecuteNonQuery();

                            MessageBox.Show("Veri başarıyla silindi!");
                        }
                    }

                    // ComboBox'tan da seçilen öğeyi kaldır
                    comboBox7.Items.RemoveAt(comboBox7.SelectedIndex);
                    MessageBox.Show("Seçilen öğe başarıyla kaldırıldı.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Hata oluştu: " + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Lütfen silinecek bir öğe seçin.");
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            // textboxtan veri alma
            string dataToInsert = textBox1.Text;

            // texboxa veri var mı diye kontrol
            if (!string.IsNullOrEmpty(dataToInsert))
            {
                try
                {
                    using (SqlConnection connection = new SqlConnection(connectionString))
                    {
                        connection.Open();

                        using (SqlCommand command = new SqlCommand("INSERT INTO ogretmen (ogr_adi_soyadi) VALUES (@Data)", connection))
                        {
                            // veri ekleme
                            command.Parameters.AddWithValue("@Data", dataToInsert);

                            command.ExecuteNonQuery();

                            MessageBox.Show("Veriler başarıyla eklendi!");
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Lütfen veri girin!");
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            string data2 = textBox2.Text;
            string data3 = textBox3.Text;
            string data4 = textBox4.Text;

            if (!string.IsNullOrEmpty(data2) && !string.IsNullOrEmpty(data3) && !string.IsNullOrEmpty(data4))
            {
                try
                {
                    using (SqlConnection connection = new SqlConnection(connectionString))
                    {
                        connection.Open();

                        // textboxtaki verileri ekleme
                        using (SqlCommand command = new SqlCommand("INSERT INTO ders (ders_kodu, ders_adi, seviye) VALUES (@Data2, @Data3, @Data4)", connection))
                        {
                            // parametre ekleme
                            command.Parameters.AddWithValue("@Data2", data2);
                            command.Parameters.AddWithValue("@Data3", data3);
                            command.Parameters.AddWithValue("@Data4", data4);

                            command.ExecuteNonQuery();

                            MessageBox.Show("Veri Tabanına Eklendi!");
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Lütfen verileri eksiksiz giriniz!");
            }
        }




        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (comboBox6.SelectedIndex != -1)
            {
                DataRowView selectedRow = (DataRowView)comboBox6.SelectedItem;
                int primaryKeyValue = (int)selectedRow["ogr_id"];

                try
                {
                    using (SqlConnection connection = new SqlConnection(connectionString))
                    {
                        connection.Open();

                        using (SqlCommand command = new SqlCommand("DELETE FROM ogretmen WHERE ogr_id = @ogr_id", connection))
                        {
                            command.Parameters.AddWithValue("@ogr_id", primaryKeyValue);
                            command.ExecuteNonQuery();

                            MessageBox.Show("Veri başarıyla silindi!");
                        }
                    }

                    // ComboBox'tan da seçilen öğeyi kaldırma
                    comboBox6.Items.RemoveAt(comboBox6.SelectedIndex);
                    MessageBox.Show("Seçilen öğe başarıyla kaldırıldı.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Hata oluştu: " + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Lütfen silinecek bir öğe seçin.");
            }
        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            string selectedDay = comboBox1.Text; // Gün
            string selectedStartTime = comboBox3.Text; // Başlama Saati
            string selectedEndTime = comboBox5.Text; // Bitiş Saati
            string selectedLesson = comboBox4.Text; // Ders Adı
            string selectedTeacher = comboBox2.Text;// Öğretmen Adı
            string selectedClass = comboBox8.Text;// Sınıf Adı
            
            string query = "SELECT dbo.aktif_ders.gün, dbo.ders.ders_adi, dbo.aktif_ders.bas_saat, dbo.aktif_ders.bit_saati, dbo.ogretmen.ogr_adi_soyadi, dbo.sinif.sinif_id FROM     dbo.aktif_ders INNER JOIN dbo.ders ON dbo.aktif_ders.ders_id = dbo.ders.ders_id INNER JOIN dbo.ogretmen ON dbo.aktif_ders.ogr_id = dbo.ogretmen.ogr_id INNER JOIN dbo.sinif ON dbo.aktif_ders.sinif_id = dbo.sinif.sinif_id WHERE  (dbo.aktif_ders.gün = @day) AND (dbo.ders.ders_adi = @lesson) AND (dbo.aktif_ders.bas_saat = @startTime) AND (dbo.aktif_ders.bit_saati = @endTime) AND (dbo.ogretmen.ogr_adi_soyadi = @teacher) AND   (dbo.sinif.sinif_id = @class)";
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    DataTable dt = new DataTable();
                    SqlDataReader reader = null;

                    connection.Open();

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        // Parametreleri sorguya ekleme
                        command.Parameters.AddWithValue("@day", selectedDay);
                        command.Parameters.AddWithValue("@startTime", selectedStartTime);
                        command.Parameters.AddWithValue("@endTime", selectedEndTime);
                        command.Parameters.AddWithValue("@lesson", selectedLesson);
                        command.Parameters.AddWithValue("@teacher", selectedTeacher);
                        command.Parameters.AddWithValue("@class", selectedClass);

                        // Çakışma kontrolü
                        reader = command.ExecuteReader();
                        dt.Load(reader);

                        // Çakışma varsa
                        if (dt.Rows .Count > 0)
                        {
                            MessageBox.Show("Çakışma var!");
                            button1.BackColor = Color.Red;
                        }
                        // Çakışma yoksa
                        else
                        {
                            MessageBox.Show("Çakışma yok!");
                            button1.Enabled = true;
                            button1.BackColor = Color.Green;
                        }
                    }
                }
                

            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata oluştu: " + ex.Message);
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string selectedDay = comboBox1.Text; // Gün
            string selectedStartTime = comboBox3.Text; // Başlama Saati
            string selectedEndTime = comboBox5.Text; // Bitiş Saati
            string selectedLesson = comboBox4.Text; // Ders Adı
            string selectedTeacher = comboBox2.Text;// Öğretmen Adı
            string selectedClass = comboBox8.Text;// Sınıf Adı

            string query = "INSERT INTO aktif_ders(gün, bas_saat, bit_saati, ders_id, ogr_id, sinif_id) VALUES  (@day,@startTime,@endTime,(SELECT ders_id FROM ders WHERE ders_adi = @lesson),(SELECT ogr_id FROM ogretmen WHERE ogr_adi_soyadi = @teacher),(SELECT sinif_id FROM sinif WHERE sinif_id = @class))"; 
                
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    DataTable dt = new DataTable();
                    SqlDataReader reader = null;

                    connection.Open();

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        // Parametreleri ekleme
                        command.Parameters.AddWithValue("@day", selectedDay);
                        command.Parameters.AddWithValue("@startTime", selectedStartTime);
                        command.Parameters.AddWithValue("@endTime", selectedEndTime);
                        command.Parameters.AddWithValue("@lesson", selectedLesson);
                        command.Parameters.AddWithValue("@teacher", selectedTeacher);
                        command.Parameters.AddWithValue("@class", selectedClass);

                        // Ders Ekleme
                        reader = command.ExecuteReader();
                        dt.Load(reader);

                        
                        if (dt.Rows.Count == 0)
                        {
                            MessageBox.Show("Ders Eklendi!");
                            button1.BackColor = Color.Orange;
                            button1.Enabled = false;
                        }
                        else
                        {
                            MessageBox.Show("Ders Eklenmedi!");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata oluştu: " + ex.Message);
            }

        }

        public void button2_Click(object sender, EventArgs e)
        {
            //Tablo güncelleme
            string query = "SELECT aktif_ders.*, ders.*, ogretmen.*, sinif.*FROM aktif_ders INNER JOIN ders ON aktif_ders.ders_id = ders.ders_id INNER JOIN ogretmen ON aktif_ders.ogr_id = ogretmen.ogr_id INNER JOIN sinif ON aktif_ders.sinif_id = sinif.sinif_id ORDER BY aktif_ders_id desc";
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    DataTable dt = new DataTable();
                    SqlDataReader reader = null;


                    connection.Open();
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                        reader = command.ExecuteReader();
                        dt.Load(reader);
                        dataGridView1.DataSource = dt;
                    


                }

                    

                }


                
            }
            catch (Exception ex)
            {

                MessageBox.Show("Hata oluştu: grid yenileme hatası " + ex.Message);
            }


        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            
           
        }

        private void button8_Click(object sender, EventArgs e)
        {
            string selectedStage = comboBox9.Text;
            string query = "SELECT aktif_ders.*, ders.*, ogretmen.*, sinif.*FROM aktif_ders INNER JOIN ders ON aktif_ders.ders_id = ders.ders_id INNER JOIN ogretmen ON aktif_ders.ogr_id = ogretmen.ogr_id INNER JOIN sinif ON aktif_ders.sinif_id = sinif.sinif_id WHERE seviye= @stage";


            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    DataTable dt = new DataTable();
                    SqlDataReader reader = null;

                    connection.Open();

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        // Parametreleri ekleme
                        command.Parameters.AddWithValue("@stage", selectedStage);
                        

                        //// Ders Ekleme
                        reader = command.ExecuteReader();
                        dt.Load(reader);
                        dataGridView1.DataSource= dt;


                        //if (dt.Rows.Count == 0)
                        //{
                        //    MessageBox.Show("Ders programı hazır!");
                        //    button1.BackColor = Color.Orange;
                        //    button1.Enabled = false;
                        //}
                        //else
                        //{
                        //    MessageBox.Show("hata!");
                        //}
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata oluştu: 352" + ex.Message);
            }



        }

        private void button9_Click(object sender, EventArgs e)
        {
            string selectedClass = comboBox10.Text;
            string query = "SELECT aktif_ders.*, ders.*, ogretmen.* FROM aktif_ders INNER JOIN ders ON aktif_ders.ders_id = ders.ders_id INNER JOIN ogretmen ON aktif_ders.ogr_id = ogretmen.ogr_id  WHERE sinif_id = @class";


            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    DataTable dt = new DataTable();
                    SqlDataReader reader = null;

                    connection.Open();

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        // Parametreleri ekleme
                        command.Parameters.AddWithValue("@class", selectedClass);


                        //// Ders Ekleme
                        reader = command.ExecuteReader();
                        dt.Load(reader);
                        dataGridView1.DataSource = dt;


                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata oluştu: 352" + ex.Message);
            }



        }
        private void ExportToExcel(DataGridView dataGridView)
        {
            if (dataGridView.Rows.Count == 0)
            {
                MessageBox.Show("Aktarılacak veri bulunamadı.");
                return;
            }

            try
            {
                Excel.Application excelApp = new Excel.Application();
                excelApp.Application.Workbooks.Add(Type.Missing);
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelApp.ActiveSheet;

                for (int i = 1; i < dataGridView.Columns.Count + 1; i++)
                {
                    excelWorksheet.Cells[1, i] = dataGridView.Columns[i - 1].HeaderText;
                }

                for (int i = 0; i < dataGridView.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView.Columns.Count; j++)
                    {
                        object cellValue = dataGridView.Rows[i].Cells[j].Value;
                        excelWorksheet.Cells[i + 2, j + 1] = cellValue != null ? cellValue.ToString() : string.Empty;

                    }
                }

                excelApp.Columns.AutoFit();
                excelApp.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Excel'e aktarım sırasında bir hata oluştu: " + ex.Message);
            }
        }

        private void ExportToPdf(DataGridView dataGridView1, string fileName)
        {
            try
            {
                // PDF belgesini oluştur
                PdfWriter writer = new PdfWriter(fileName);
                PdfDocument pdf = new PdfDocument(writer);
                Document document = new Document(pdf);

            
                // DataGridView içeriğini PDF'e ekle
                Table table = new Table(dataGridView1.ColumnCount);
                for (int i = 0; i < dataGridView1.ColumnCount; i++)
                {
                    table.AddCell(new Cell().Add(new Paragraph(dataGridView1.Columns[i].HeaderText)));
                }

                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    {
                        table.AddCell(new Cell().Add(new Paragraph(dataGridView1[j, i].Value.ToString())));
                    }
                }

                document.Add(table);

                // PDF dosyasını kapat
                document.Close();

                MessageBox.Show("PDF dosyası başarıyla oluşturuldu.", "Başarı", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata oluştu: " + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        
    }

        private void button10_Click(object sender, EventArgs e)
        {

            string selectedDay = comboBox11.Text;
            string query = "SELECT aktif_ders.*, ders.*, ogretmen.* FROM aktif_ders INNER JOIN ders ON aktif_ders.ders_id = ders.ders_id INNER JOIN ogretmen ON aktif_ders.ogr_id = ogretmen.ogr_id  WHERE gün = @day";


            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    DataTable dt = new DataTable();
                    SqlDataReader reader = null;

                    connection.Open();

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        // Parametreleri ekleme
                        command.Parameters.AddWithValue("@day", selectedDay);


                        //// Ders Ekleme
                        reader = command.ExecuteReader();
                        dt.Load(reader);
                        dataGridView1.DataSource = dt;


                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata oluştu: 352" + ex.Message);
            }


        }

        private void button11_Click(object sender, EventArgs e)
        {
            ExportToPdf(dataGridView1, "DataGridViewToPdf.pdf");
        }

        private void button12_Click(object sender, EventArgs e)
        {
            ExportToExcel(dataGridView1);
        }
    } }
  
        
