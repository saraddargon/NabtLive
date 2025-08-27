using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Telerik.WinControls.UI;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Data.SqlClient;
namespace StockControl
{
    public static class dbShowData2
    {
        public static void InsertToExcel(ref Excel.Worksheet exc, string Column, dynamic Values)
        {
            try
            {
                Excel.Range refs = exc.get_Range(Column);
                refs.Value2 = Values;
                if (Values.Contains("Æ"))
                {
                    int addint = Values.IndexOf("Æ");
                    refs.Characters[addint, 2].Font.Name = "Symbol";//                          
                }
            }
            catch { }
        }
        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                // MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        private static string Getcolumn(int Col)
        {
            string RT = "";

            if (Col.Equals(1))
                RT = "A";
            else if (Col.Equals(2))
                RT = "B";
            else if (Col.Equals(3))
                RT = "C";
            else if (Col.Equals(4))
                RT = "D";
            else if (Col.Equals(5))
                RT = "E";
            else if (Col.Equals(6))
                RT = "F";
            else if (Col.Equals(7))
                RT = "G";
            else if (Col.Equals(8))
                RT = "H";
            else if (Col.Equals(9))
                RT = "I";
            else if (Col.Equals(10))
                RT = "J";
            else if (Col.Equals(11))
                RT = "K";
            else if (Col.Equals(12))
                RT = "L";
            else if (Col.Equals(13))
                RT = "M";
            else if (Col.Equals(14))
                RT = "N";
            else if (Col.Equals(15))
                RT = "O";
            else if (Col.Equals(16))
                RT = "P";
            else if (Col.Equals(17))
                RT = "Q";
            else if (Col.Equals(18))
                RT = "R";
            else if (Col.Equals(19))
                RT = "S";
            else if (Col.Equals(20))
                RT = "T";
            else if (Col.Equals(21))
                RT = "U";
            else if (Col.Equals(22))
                RT = "V";
            else if (Col.Equals(23))
                RT = "W";
            else if (Col.Equals(24))
                RT = "X";
            else if (Col.Equals(25))
                RT = "Y";
            else if (Col.Equals(26))
                RT = "Z";

            else if (Col.Equals(27))
                RT = "AA";
            else if (Col.Equals(28))
                RT = "AB";
            else if (Col.Equals(29))
                RT = "AC";
            else if (Col.Equals(30))
                RT = "AD";
            else if (Col.Equals(31))
                RT = "AE";
            else if (Col.Equals(32))
                RT = "AF";
            else if (Col.Equals(33))
                RT = "AG";
            else if (Col.Equals(34))
                RT = "AH";
            else if (Col.Equals(35))
                RT = "AI";
            else if (Col.Equals(36))
                RT = "AJ";
            else if (Col.Equals(37))
                RT = "AK";
            else if (Col.Equals(38))
                RT = "AL";
            else if (Col.Equals(39))
                RT = "AM";
            else if (Col.Equals(40))
                RT = "AN";
            else if (Col.Equals(41))
                RT = "AO";
            else if (Col.Equals(42))
                RT = "AP";
            else if (Col.Equals(43))
                RT = "AQ";

            else if (Col.Equals(44))
                RT = "AR";
            else if (Col.Equals(45))
                RT = "AS";
            else if (Col.Equals(46))
                RT = "AT";
            else if (Col.Equals(47))
                RT = "AU";
            else if (Col.Equals(48))
                RT = "AV";
            else if (Col.Equals(49))
                RT = "AW";
            else if (Col.Equals(50))
                RT = "AX";


            return RT;
        }
        public static void LoadToTempVersion(string QCNo)
        {
            using (DataClasses1DataContext db = new DataClasses1DataContext())
            {
                tb_QCHD qh = db.tb_QCHDs.Where(p => p.QCNo.Equals(QCNo)).FirstOrDefault();
                if (qh != null)
                {
                    tb_ProductionHD pd = db.tb_ProductionHDs.Where(q => q.OrderNo.Equals(qh.WONo)).FirstOrDefault();
                    if (pd != null)
                    {
                        db.sp_49_1_QC_LoadToTemp(pd.PartFG, qh.FormISO, pd.Createdate, dbClss.UserID);
                    }
                }
            }
        }
        public static void DeleteUserPrintQC()
        {
            try
            {
                //Delete Data TempTable Header//
                using (SqlConnection connection = new SqlConnection(dbClss.DbConn))
                {
                    string QueryD = "delete from [dbo].[tb_QCPrintHeader] where UserID='" + dbClss.UserID + "'";
                    SqlCommand command = new SqlCommand(QueryD, connection);
                    command.ExecuteNonQuery();
                }
            }
            catch { }
        }
        public static void PrintFMQC055_NewA1(string WO, string PartNo, string QCNo1)
        {
            //11/aug/23 Create
            try
            {
                //Step Report 055

                string DATA = AppDomain.CurrentDomain.BaseDirectory;
                string tempPath = System.IO.Path.GetTempPath();
                string FileName = "FM-QA-055R.04.xlsx";
                string tempfile = tempPath + FileName;
                DATA = DATA + @"QC\" + FileName;
                if (File.Exists(tempfile))
                {
                    try
                    {
                        File.Delete(tempfile);
                    }
                    catch { }
                }
                //Load Version//
                LoadToTempVersion(QCNo1);//Load Version
               // DeleteUserPrintQC(); //Delete User

                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook excelBook = excelApp.Workbooks.Open(
                  DATA, 0, true, 5,
                  "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false,
                  0, true);
                Excel.Sheets sheets = excelBook.Worksheets;

                ///////////Query Data//
                string Data1 = "";
                int QtyTAG = 0;
                bool Page1 = true;
                bool Page2 = false;
                bool Page3 = false;
                bool Page4 = false;
                bool Page5 = false;
                bool Page6 = false;

                using (SqlConnection connection = new SqlConnection(dbClss.DbConn))
                {
                    string query = "SELECT top 1 OfTAG FROM tb_QCTAG where QCNo='" + QCNo1 + "'";
                    SqlCommand command = new SqlCommand(query, connection);
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Data1 = reader["OfTAG"].ToString();
                            string[] Data = Data1.Split(',');
                            string[] PPTAG2 = Data[5].ToLower().Split('f');
                            QtyTAG = Convert.ToInt32(PPTAG2[1]);
                        }
                    }
                    if (QtyTAG > 0)
                    {
                        //Call Store//
                        string query2 = "sp_90_ExcelHeader";
                        SqlCommand command2 = new SqlCommand(query2, connection);
                        command2.CommandType = System.Data.CommandType.StoredProcedure;  
                        command2.Parameters.AddWithValue("@UserID", dbClss.UserID);                       
                        command2.Parameters.AddWithValue("@WONo", WO);
                        command2.Parameters.AddWithValue("@QCNo", QCNo1);
                        command2.Parameters.AddWithValue("@PartNo", PartNo);                        
                        command2.Parameters.AddWithValue("@MaxTag", QtyTAG);
                        command2.ExecuteNonQuery();
                        //Call Store Line//


                        if (QtyTAG>0)
                        {
                            Page1 = true;                            
                            Excel.Worksheet worksheet = (Excel.Worksheet)sheets.get_Item(1);
                        }
                        if (QtyTAG > 25)
                        {
                            Page2 = true;
                            Excel.Worksheet worksheet2 = (Excel.Worksheet)sheets.get_Item(2);
                        }
                        if (QtyTAG > 50)
                        {
                            Page3 = true;
                            Excel.Worksheet worksheet3 = (Excel.Worksheet)sheets.get_Item(3);
                        }
                        if (QtyTAG > 75)
                        {
                            Page4 = true;
                            Excel.Worksheet worksheet4 = (Excel.Worksheet)sheets.get_Item(4);
                        }
                        if (QtyTAG > 100)
                        {
                            Page5 = true;
                            Excel.Worksheet worksheet5 = (Excel.Worksheet)sheets.get_Item(5);
                        }
                        if (QtyTAG > 125)
                        {
                            Page6 = true;
                            Excel.Worksheet worksheet6 = (Excel.Worksheet)sheets.get_Item(6);
                        }
                    }//QtyTAG Close


                    //Close Tag Connection//
                }

                excelBook.SaveAs(tempfile);
                excelBook.Close(false);
                excelApp.Quit();              
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                GC.GetTotalMemory(false);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.GetTotalMemory(true);
                System.Diagnostics.Process.Start(tempfile);

            }
            catch { }

        } //Public

    }
}
