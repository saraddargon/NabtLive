using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Linq;
using Microsoft.VisualBasic.FileIO;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace StockControl
{
    public partial class UploadSetDataExport : Telerik.WinControls.UI.RadRibbonForm
    {
   
        Telerik.WinControls.UI.RadTextBox CodeNo_tt = new Telerik.WinControls.UI.RadTextBox();
        int screen = 0;
        public UploadSetDataExport(Telerik.WinControls.UI.RadTextBox  CodeNox)
        {
            this.Name = "UploadSetDataExport";
            InitializeComponent();
            CodeNo_tt = CodeNox;
            screen = 1;
        }
        public UploadSetDataExport()
        {
            this.Name = "UploadSetDataExport";
            InitializeComponent();
        }

        string PR1 = "";
        string PR2 = "";
        string Type = "";
        //private int RowView = 50;
        //private int ColView = 10;
        //DataTable dt = new DataTable();
        private void radMenuItem2_Click(object sender, EventArgs e)
        {

        }

        private void radRibbonBar1_Click(object sender, EventArgs e)
        {

        }
        private void GETDTRow()
        {
            //dt.Columns.Add(new DataColumn("UnitCode", typeof(string)));
            //dt.Columns.Add(new DataColumn("UnitDetail", typeof(string)));
            //dt.Columns.Add(new DataColumn("UnitActive", typeof(bool)));
        }
        private void Unit_Load(object sender, EventArgs e)
        {
            getDT();
        }
        DataTable dt_d = new DataTable();
        private void getDT()
        {
            dt_d = new DataTable();
            dt_d.Columns.Add(new DataColumn("PartNapt", typeof(string)));
            dt_d.Columns.Add(new DataColumn("PartCust", typeof(string)));
            dt_d.Columns.Add(new DataColumn("Desc", typeof(string)));
            dt_d.Columns.Add(new DataColumn("ThaiLang", typeof(string)));
            dt_d.Columns.Add(new DataColumn("Point", typeof(string)));
            dt_d.Columns.Add(new DataColumn("Fomula", typeof(string)));

        }
        private void radButton1_Click(object sender, EventArgs e)
        {
            //openFileDialog1.Filter = "excel files (*.xlsx)|*.xlsx";
            //openFileDialog1.FilterIndex = 2;
            //openFileDialog1.RestoreDirectory = true;
            //openFileDialog1.FileName = "";

            try
            {
                this.Cursor = Cursors.WaitCursor;
                txtPartFile.Text = "";
                dt_d.Rows.Clear();
                openFileDialog1.FilterIndex = 2;
                openFileDialog1.RestoreDirectory = true;

                openFileDialog1.DefaultExt = "*.xls";
                openFileDialog1.AddExtension = true;
                openFileDialog1.FileName = "";
                openFileDialog1.Filter = "Excel 2003-2010  (*.xls,*.xlsx,*.csv)|*.xls;*xlsx;*.csv";

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    txtPartFile.Text = openFileDialog1.FileName;
                    //string name = Path.GetFileName(openFileDialog1.FileName);
                    string Exten = Path.GetExtension(openFileDialog1.FileName);
                    if (Exten.ToUpper() == ".XLS" || Exten.ToUpper() == ".XLSX")
                        Import_Excel(openFileDialog1.FileName);
                    //else if (Exten.ToUpper() == ".CSV")
                    //    Import_CSV(openFileDialog1.FileName);

                    if (dt_d.Rows.Count > 0)
                        lblSS.Visible = true;
                    else
                    {
                        MessageBox.Show("can't load data import.");
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            finally { this.Cursor = Cursors.Default; }
        }
        private void Import_Excel(string Name)
        {
            try
            {
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook theWorkbook = excelApp.Workbooks.Open(
                  openFileDialog1.FileName, 0, true, 5,
                  "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false,
                  0, true);


                Excel.Sheets sheets = theWorkbook.Worksheets;
                Excel.Worksheet worksheet = (Excel.Worksheet)sheets.get_Item(1);

                progressBar1.Visible = true;
                progressBar1.Maximum = 10003;
                progressBar1.Minimum = 1;
                int icount = 0;

                //int Sheet4 = 0;
                for (int j = 2; j < 1003; j++)
                {
                    if (progressBar1.Value < progressBar1.Maximum)
                    {
                        progressBar1.Value = icount + 1;
                        icount = icount + 1;
                        progressBar1.PerformStep();
                    }

                    System.Array myvalues;
                    Excel.Range range = worksheet.get_Range("A" + j.ToString(), "F" + j.ToString());
                    myvalues = (System.Array)range.Cells.Value;
                    string[] strArray = ConvertToStringArray(myvalues);
                    if (!Convert.ToString(strArray[0]).Equals("")
                        //!Convert.ToString(strArray[2]).Equals("")
                        )
                    {
                        GetDataSystem2(Convert.ToString(strArray[0]).Trim() //Code
                            , Convert.ToString(strArray[1]).Trim()//CustItemName
                            , Convert.ToString(strArray[2]).Trim()//CustItemNo
                            , Convert.ToString(strArray[3]).Trim()//CustoemrName
                            , Convert.ToString(strArray[4]).Trim()//CustomerNo  
                            , Convert.ToString(strArray[5]).Trim()//CustomerNo  
                            );
                    }
                    else
                        break;
                }
                progressBar1.Value = progressBar1.Maximum;
                progressBar1.PerformStep();
                progressBar1.Visible = false;

                //excelBook.Save();
                //excelApp.Quit();
                releaseObject(worksheet);

                releaseObject(excelApp);
                //Marshal.FinalReleaseComObject(worksheet);


            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }

        }
        int RowIndex_temp = 0;
        private void GetDataSystem2(string PartNapt,string PartCust, string desc, string thailang, string point,string Fomula
           )
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;

                RowIndex_temp = (dt_d.Rows.Count);
                DataRow dr = dt_d.NewRow();

                string Status = "OK";

                dr["PartNapt"] = PartNapt;
              
                //if (dbClss.TSt(CustItemNo).Trim() != "" && Status == "OK")
                    dr["PartCust"] = PartCust;
                //else
                //    Status = "NG";
                //if (dbClss.TSt(CustItemName).Trim() != "" && Status == "OK")
                    dr["Desc"] = desc;
                //else
                //    Status = "NG";

                dr["ThaiLang"] = thailang;
                dr["Point"] = point;
                dr["Fomula"] = Fomula;



                if (Status == "OK")
                {
                    dt_d.Rows.Add(dr);
                    RowIndex_temp += 1;
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            //finally { this.Cursor = Cursors.Default; }
        }

        private void Import_CSV(string Name)
        {
            using (TextFieldParser parser = new TextFieldParser(Name, Encoding.GetEncoding("windows-874")))
            {
                this.Cursor = Cursors.WaitCursor;

                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");
                int a = 0;
                int c = 0;

                string Code, CustItemNo, CustItemName, CustomerNo = "";
                string CSTMName = "";
                while (!parser.EndOfData)
                {
                    //Processing row
                    a += 1;
                    Code = ""; CustItemNo = ""; CustItemName = ""; CustomerNo = "";
                  
                    string[] fields = parser.ReadFields();
                    c = 0;
                    foreach (string field in fields)
                    {
                        c += 1;
                        ////TODO: Process field
                        //    // MessageBox.Show(field);
                        if (a >= 2)
                        {
                            if (c == 1 && Convert.ToString(field).Equals(""))
                            {
                                break;
                            }

                            if (c == 1)
                                Code = Convert.ToString(field);
                            else if (c == 2)
                                CustItemName = Convert.ToString(field);
                            else if (c == 3)
                                CustItemNo = Convert.ToString(field);                            
                            else if (c == 4)
                                CSTMName = Convert.ToString(field);
                            else if (c == 5)
                                CustomerNo = Convert.ToString(field);
                         
                        }

                    }
                    if (Code != "")
                    {
                        GetDataSystem2(Code,  CustItemName, CustItemNo,CSTMName, CustomerNo,"");
                    }

                }

            }
        }
        private string[] ConvertToStringArray(System.Array values)
        {

            // create a new string array
            string[] theArray = new string[values.Length];

            // loop through the 2-D System.Array and populate the 1-D String Array
            for (int i = 1; i <= values.Length; i++)
            {
                if (values.GetValue(1, i) == null)
                    theArray[i - 1] = "";
                else
                    theArray[i - 1] = (string)values.GetValue(1, i).ToString();
            }

            return theArray;
        }
        private void releaseObject(object obj)
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
        private void btnExport_Click(object sender, EventArgs e)
        {
            //Upload

            try
            {
                if (MessageBox.Show("Update Data?", "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {


                    // tb_CheckStockList <= Insert to this table
                    //Update Status tb_CheckStock to "Waiting Check"
                    //สามารถอัพโหลดใหม่ได้ โดยการ ให้ลบ ข้อมูลเก่าทั้งหมดออกก่อน

                    //    string DKUBU, ItemCode, ItemDescription, Type
                    //, Revision, ExclusionClass, StorageWorkCenter, StorageWorkCenterName
                    //, CurrentInventory, InventoryValue, StockBeforeInventory, PhysicalInventoryValue
                    //, UnitOfMeasure = "";

                    int C = 0;
                    this.Cursor = Cursors.WaitCursor;
                    using (DataClasses1DataContext db = new DataClasses1DataContext())
                    {
                        //radProgressBarElement1.Visibility = Telerik.WinControls.ElementVisibility.Visible;
                        //radProgressBarElement1.Minimum = 0;
                        //radProgressBarElement1.Maximum = dt_d.Rows.Count;

                        foreach (DataRow dr in dt_d.Rows)
                        {
                            //dbClss.TSt(dr["Code"]      
                            C += 1;
                            tb_InvoiceExMasterMap ivc = db.tb_InvoiceExMasterMaps.Where(p => p.PartNapt.Equals(dbClss.TSt(dr["PartNapt"]))).FirstOrDefault();
                            if (ivc != null)
                            {
                                ivc.PartCustomer = dbClss.TSt(dr["PartCust"]);
                                ivc.Description = dbClss.TSt(dr["Desc"]);
                                ivc.Point = dbClss.TSt(dr["Point"]);
                                ivc.ThaiLang = dbClss.TSt(dr["ThaiLang"]);
                                ivc.FomulaA = dbClss.TSt(dr["Fomula"]);
                                db.SubmitChanges();
                            }
                            else
                            {
                                tb_InvoiceExMasterMap ivn = new tb_InvoiceExMasterMap();
                                ivn.PartNapt = dbClss.TSt(dr["PartNapt"]);
                                ivn.PartCustomer = dbClss.TSt(dr["PartCust"]);
                                ivn.Description = dbClss.TSt(dr["Desc"]);
                                ivn.Point = dbClss.TSt(dr["Point"]);
                                ivn.ThaiLang = dbClss.TSt(dr["ThaiLang"]);
                                ivn.FomulaA = dbClss.TSt(dr["Fomula"]);
                                db.tb_InvoiceExMasterMaps.InsertOnSubmit(ivn);
                                db.SubmitChanges();
                            }                                
                        }

                        if (C > 0)
                        {
                            MessageBox.Show("Import data Complete.");
                        }
                        else
                        {
                            MessageBox.Show("ไม่พบข้อมูล!");
                        }

                        //radProgressBarElement1.Visibility = Telerik.WinControls.ElementVisibility.Collapsed;

                    }
                    lblSS.Visible = false;
                    txtPartFile.Text = "";
                }
                else
                {
                    return;
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            finally { this.Cursor = Cursors.Default; }
        }

        private void radButtonElement1_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            if (MessageBox.Show("คุณต้องการลบข้อมูลทั้งหมด หรือไม่ ?", "ลบรายการ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    var mp = db.tb_InvoiceExMasterMaps.ToList();
                    if(mp.Count>0)
                    {
                        db.tb_InvoiceExMasterMaps.DeleteAllOnSubmit(mp);
                        db.SubmitChanges();
                        MessageBox.Show("ลบข้อมูลเรียบร้อย");
                    }
                }
            }
            this.Cursor = Cursors.Default ;

        }
    }
}
