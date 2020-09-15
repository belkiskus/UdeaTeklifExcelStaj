using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDataReader;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace ExcelStajProjesi
{
    public partial class Form1 : Form
    {

        List<Teklif> TeklifList = new List<Teklif>();
        DataTable dt;
        public Form1()
        {
            InitializeComponent();
        }


        //Excel'den okunan verileri Datatable'a aktarma

        DataTableCollection dTable;
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog excelfile = new OpenFileDialog() 
            { Filter = "Excel Dosyaları|*.xlsx|Excel Dosyaları|*.xls", ValidateNames = true };   //Openfiledialog ile excel dosyaları açıldı
            
                if(excelfile.ShowDialog() == DialogResult.OK)                                    //Ok basıldığında
                {                                                                                
                    textBox1.Text = excelfile.FileName;                                         //textboxa dosya uzantısı yazdırıldı
                    FileStream readexcel = File.Open(excelfile.FileName, FileMode.Open, FileAccess.Read);         //Filestream ile readexcel oluşturuldu ve excel dosyası açıldı
                    IExcelDataReader reader = ExcelReaderFactory.CreateReader(readexcel);                         //Datareader ile readexcel okunarak reader oluşturuldu

                    DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable=(x)=>new ExcelDataTableConfiguration() { UseHeaderRow=true}
                    });

                    dTable = result.Tables;

                    comboBox1.Items.Clear();
                    foreach (DataTable item in dTable)
                    {
                    comboBox1.Items.Add(item.TableName);
                    }
                    reader.Close();

                    
                }

        }

        //Comboboxtan seçilen Excel sayfasını Datatable'dan Datagridview'e aktarma
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

            dt = dTable[comboBox1.SelectedIndex];
            dataGridView1.DataSource = dt;

            TeklifList = Teklif.teklif_listesi_olustur(dt);

            

            //Listeden Datagirdview' e aktarma
            //dataGridView1.DataSource = Teklif.teklif_listesi_olustur(dt);

        }

        //Sayfayı Güncelle butonu
        private void button2_Click(object sender, EventArgs e)
        {
            //Datagridview'deki verileri Datatable'a aktarır
            DataTable dt = new DataTable();
            foreach (DataGridViewColumn col in dataGridView1.Columns)
            {
                dt.Columns.Add(col.Name);
            }

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                DataRow dRow = dt.NewRow();
                foreach (DataGridViewCell cell in row.Cells)
                {
                    dRow[cell.ColumnIndex] = cell.Value;
                }
                dt.Rows.Add(dRow);
            }

            //Listeyi güncelleme

            TeklifList = Teklif.teklif_listesi_olustur(dt);




        }

        //Datagridview'den Excel'e verileri aktarma
        private void button3_Click(object sender, EventArgs e)
        {
            // Excel objesi üretildi

            Microsoft.Office.Interop.Excel._Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = excel.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            //excel.Visible = true;

            try
            {

                worksheet = workbook.ActiveSheet;

                worksheet.Name = "Teklif Takip Güncel";
                worksheet.Range["A1:R1"].EntireRow.Interior.Color = System.Drawing.Color.Yellow;

                int cellrowindex = 1;
                int cellcolumnindex = 1;

                //her satırda döngü oluşturulur ve her sütundaki değer okunur

                for (int i = 0; i <= dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        //excel indeksi 1,1'den başlar. ilk satır sütun başlıklarına sahip olduğundan, bir koşul kontrolü eklenir
                        
                        if (cellrowindex == 1)
                        {
                            worksheet.Columns.AutoFit();
                            worksheet.Cells[cellrowindex, cellcolumnindex] = dataGridView1.Columns[j].HeaderText;

                        }
                        else
                        {
                            worksheet.Columns.AutoFit();
                            worksheet.Cells[cellrowindex, cellcolumnindex] = dataGridView1.Rows[i - 1].Cells[j].Value.ToString();
                        }
                        cellcolumnindex++;
                    }
                    cellcolumnindex = 1;
                    cellrowindex++;
                }

                //kullanıcıdan kaydedilecek Excel'in konumu ve dosya adı alınır

                SaveFileDialog savedialog = new SaveFileDialog();
                savedialog.Filter = "Excel Dosyaları|*.xlsx|Excel Dosyaları|*.xls";
                //savedialog.FilterIndex = 0;
                savedialog.FileName = "Udea Teklif Takip Yeni"; 

                if (savedialog.ShowDialog() == DialogResult.OK)
                {
                    workbook.SaveAs(savedialog.FileName);
                    //MessageBox.Show("Bilgiler yeni bir excel sayfasına aktarıldı.","Bilgilendirme", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                excel.Quit();
                workbook = null;
                excel = null;
            }



        }

        private void button4_Click(object sender, EventArgs e)
        {
            //Satır Sil butonu


            try
            {
                //Liste güncellenir
                TeklifList = Teklif.teklif_listesi_olustur(dt);

                if (dataGridView1.SelectedRows.Count != 0)
                {
                    foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                    {
                        dataGridView1.Rows.Remove(row);
                    }

                    //Datatable güncellenir
                    dt.AcceptChanges();

   
                }
                else
                {
                    MessageBox.Show("Lütfen silinecek satırı seçiniz.", "Bilgilendirme", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (NullReferenceException exception) {
                MessageBox.Show("Satır seçilmedi");
                
            }
            catch(System.Exception exception)
            {
                MessageBox.Show("Bir hata meydana geldi.ERROR="+exception.Message);
            }
            


        }



        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Toplamda kaç teklif verildi?
            if (comboBox2.SelectedIndex == 0)
            {
                var toplam_teklif_sayisi = TeklifList.Select(x => x.teklif_no).Count();
                textBox2.Text = "Toplam teklif sayısı :" + Environment.NewLine + Convert.ToString(toplam_teklif_sayisi);
            }

            //Toplam teklif sayısı(ind ve rev hariç)
            if (comboBox2.SelectedIndex == 1)
            {
                var istenilen_teklif_sayisi = TeklifList.Where(x => x.siparis_durumu == "yok").Count();
                textBox2.Text = "Toplam teklif sayısı (indirim ve revizyon hariç) :" + Environment.NewLine + Convert.ToString(istenilen_teklif_sayisi);
            }

            //Tekliflerin ne kadarı USD ile verildi?
            else if (comboBox2.SelectedIndex == 2)
            {
                //double USD_fiyat_toplami = TeklifList.Sum(x => x.siparis_tutari);
                int USD_fiyat_toplami = TeklifList.Sum(x => x.USD_fiyatı);               
                textBox2.Text = "Toplam teklif fiyatı(USD) :" + Environment.NewLine + Convert.ToString(USD_fiyat_toplami);
            }

            //Tekliflerin ne kadarı TL ile verildi?
            else if (comboBox2.SelectedIndex == 3)
            {
                int TL_fiyat_toplami = TeklifList.Sum(x => x.TL_fiyatı);
                textBox2.Text = "Toplam teklif fiyatı(TL) : " + Environment.NewLine + Convert.ToString(TL_fiyat_toplami);
            }

            //Toplamdan kaç kez indirim yapıldı?
            else if (comboBox2.SelectedIndex == 4)
            {
                var indirim_yapilan_teklif_sayisi = TeklifList.Where(x => x.indirim == "indirim").Count();
                textBox2.Text = "İndirim yapılan teklif sayısı :" + Environment.NewLine + Convert.ToString(indirim_yapilan_teklif_sayisi);
            }

            //Toplamda kaç USD teklif verildi?
            else if (comboBox2.SelectedIndex == 5)
            {
                var USD_teklif_sayisi = TeklifList.Where(x => x.para_birimi == "USD").Count();
                textBox2.Text = "Teklif sayısı(USD) :" + Environment.NewLine + Convert.ToString(USD_teklif_sayisi);
            }

            //Toplamda kaç USD sipariş açıldı?
            else if (comboBox2.SelectedIndex == 6)
            {
                int USD_ile_acilan_teklif_sayisi = TeklifList.Where(x => x.siparis_durumu == "Sipariş Açıldı").Sum(x => x.USD_fiyatı);
                textBox2.Text = "Toplam sipariş fiyatı(USD) :" + Environment.NewLine + Convert.ToString(USD_ile_acilan_teklif_sayisi);
            }

            ////Toplamda kaç TL teklif verildi?
            else if (comboBox2.SelectedIndex == 7)
            {
                var TL_teklif_sayisi = TeklifList.Where(x => x.para_birimi == "TL").Count();
                textBox2.Text = "Teklif sayısı(TL) :" + Environment.NewLine + Convert.ToString(TL_teklif_sayisi);
            }

            //Toplamda kaç TL sipariş açıldı?
            else if (comboBox2.SelectedIndex == 8)
            {
                int TL_ile_acilan_siparis_fiyati = TeklifList.Where(x => x.siparis_durumu == "Sipariş Açıldı").Sum(x => x.TL_fiyatı);
                textBox2.Text = "Toplam sipariş fiyatı(TL) :" + Environment.NewLine + Convert.ToString(TL_ile_acilan_siparis_fiyati);
            }

            else
                return;
        }



        //Satırın üzerine çift tıklandığında, tıklanan satırın ardına yeni bir satır eklenir
        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            DataRow yeni_satir = dt.NewRow();
            int cift_tiklanan_satirin_indexi = dataGridView1.CurrentRow.Cells[0].RowIndex;
            //MessageBox.Show(cift_tiklanan_satirin_indexi.ToString()+ Environment.NewLine);
            dt.Rows.InsertAt(yeni_satir, (cift_tiklanan_satirin_indexi + 1));

            //Datatable güncellenir
            dt.AcceptChanges();

        }


        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            //string teklif_adi = textBox3.Text;
            int proje_sayisi = TeklifList.Where(x => x.proje == textBox3.Text).ToList().Count;
            double indirim_yuzde_toplami = TeklifList.Where(x => x.proje == textBox3.Text &&  x.indirim == "indirim").Sum(x => x.indirim_yuzdesi);
            //textBox2.Text = Convert.ToString(proje_sayisi) + " tane indirim bulundu." + Environment.NewLine + "İndirim yüzde toplamı : " + Convert.ToString(indirim_yuzde_toplami);
            MessageBox.Show(Convert.ToString(proje_sayisi) + " tane indirim bulundu.           " + Environment.NewLine + "İndirim yüzde toplamı : " + Convert.ToString(indirim_yuzde_toplami),"Yüzde Toplamı");

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }
    }
}

