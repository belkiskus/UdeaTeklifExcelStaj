using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DataTable = System.Data.DataTable;

namespace ExcelStajProjesi
{

    class Teklif
    {
        public int teklif_no { get; set; }
        public string teklif_verilen_kurum { get; set; }
        public string proje { get; set; }
        public string indirim { get; set; }
        public double indirim_yuzdesi { get; set; }
        public string aciklama { get; set; }
        public DateTime teklif_tarihi { get; set; }
        public string tekliflendirilen_miktar { get; set; }
        public double tutar { get; set; }
        public string para_birimi { get; set; }
        public int USD_fiyatı { get; set; }
        public int TL_fiyatı { get; set; }
        public string siparis_durumu { get; set; }
        public string siparis_no { get; set; }
        public DateTime siparis_tarihi { get; set; }
        public double siparis_tutari { get; set; }
        public string birim { get; set; }
        public string siparis_miktarı { get; set; }


        // Teklif Olusturma metodu
        public static List<Teklif> teklif_listesi_olustur(DataTable dTable)
        {
            List<Teklif> teklifList = new List<Teklif>();

            try
            {
                for (int i = 0; i < dTable.Rows.Count; i++)
                {
                    Teklif teklif = new Teklif();

                    try
                    {

                        teklif.teklif_no = dTable.Rows[i][0] != DBNull.Value ? Convert.ToInt32(dTable.Rows[i][0]) : 0;
                        teklif.teklif_verilen_kurum = dTable.Rows[i][1] != DBNull.Value ? Convert.ToString(dTable.Rows[i][1]) : "yok";
                        teklif.proje = dTable.Rows[i][2] != DBNull.Value ? Convert.ToString(dTable.Rows[i][2]) : "yok";
                        teklif.proje = teklif.proje.TrimEnd();
                        teklif.indirim = dTable.Rows[i][3] != DBNull.Value ? Convert.ToString(dTable.Rows[i][3]) : "yok";
                        teklif.indirim_yuzdesi = dTable.Rows[i][4] != DBNull.Value ? Convert.ToDouble(dTable.Rows[i][4]) : 0;
                        teklif.aciklama = dTable.Rows[i][5] != DBNull.Value ? Convert.ToString(dTable.Rows[i][5]) : "yok";
                        teklif.teklif_tarihi = dTable.Rows[i][6] != DBNull.Value ? Convert.ToDateTime(dTable.Rows[i][6]) : DateTime.Now;
                        teklif.tekliflendirilen_miktar = dTable.Rows[i][7] != DBNull.Value ? Convert.ToString(dTable.Rows[i][7]) : "yok";
                        teklif.tutar = dTable.Rows[i][8] != DBNull.Value ? Convert.ToDouble(dTable.Rows[i][8]) : 0;
                        teklif.para_birimi = dTable.Rows[i][9] != DBNull.Value ? Convert.ToString(dTable.Rows[i][9]) : "yok";
                        teklif.USD_fiyatı = dTable.Rows[i][10] != DBNull.Value ? Convert.ToInt32(dTable.Rows[i][10]) : 0;
                        teklif.TL_fiyatı = dTable.Rows[i][11] != DBNull.Value ? Convert.ToInt32(dTable.Rows[i][11]) : 0;
                        teklif.siparis_durumu = dTable.Rows[i][12] != DBNull.Value ? Convert.ToString(dTable.Rows[i][12]) : "yok";
                        teklif.siparis_no = dTable.Rows[i][13] != DBNull.Value ? Convert.ToString(dTable.Rows[i][13]) : "yok";
                        teklif.siparis_tarihi = dTable.Rows[i][14] != DBNull.Value ? Convert.ToDateTime(dTable.Rows[i][14]) : DateTime.Now;
                        teklif.siparis_tutari = dTable.Rows[i][15] != DBNull.Value ? Convert.ToDouble(dTable.Rows[i][15]) : 0;
                        teklif.birim = dTable.Rows[i][16] != DBNull.Value ? Convert.ToString(dTable.Rows[i][16]) : "yok";
                        teklif.siparis_miktarı = dTable.Rows[i][17] != DBNull.Value ? Convert.ToString(dTable.Rows[i][17]) : "yok";

                    }
                    catch (SystemException exception)
                    {
                        Console.WriteLine("Bir hata meydana geldi. ERROR=" + exception.ToString());
                    }


                    teklifList.Add(teklif);


                }
            }
            catch (Exception exception)
            {

                throw exception;
            }


            return teklifList;
        }


    }

   
}
