using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Linq;
using Microsoft.VisualBasic.FileIO;
using Telerik.WinControls;
using Telerik.WinControls.UI;

namespace StockControl
{
    public partial class WarehouseCheckList : Telerik.WinControls.UI.RadRibbonForm
    {
        public WarehouseCheckList()
        {
            InitializeComponent();
            dtDate1.Value = DateTime.Now;
            dtDate2.Value = DateTime.Now;
        }
        DataTable dt = new DataTable();
        private void radMenuItem2_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            HistoryView hw = new HistoryView(this.Name);
            this.Cursor = Cursors.Default;
            hw.ShowDialog();
        }

        private void radRibbonBar1_Click(object sender, EventArgs e)
        {

        }
        private void GETDTRow()
        {
            dt.Columns.Add(new DataColumn("edit", typeof(bool)));
            dt.Columns.Add(new DataColumn("code", typeof(string)));
            dt.Columns.Add(new DataColumn("Name", typeof(string)));
            dt.Columns.Add(new DataColumn("Active", typeof(bool)));
            dt.Columns.Add(new DataColumn("CreateDate", typeof(DateTime)));
            dt.Columns.Add(new DataColumn("CreateBy", typeof(string)));
        }
        private void Unit_Load(object sender, EventArgs e)
        {
            //RMenu3.Click += RMenu3_Click;
            //RMenu4.Click += RMenu4_Click;
            //RMenu5.Click += RMenu5_Click;
            //RMenu6.Click += RMenu6_Click;
            // radGridView1.ReadOnly = true;
            dtDate1.Value = DateTime.Now;
            dtDate2.Value = DateTime.Now;
            radGridView1.AutoGenerateColumns = false;
            GETDTRow();          
       
            DataLoad2();
        }

        private void RMenu6_Click(object sender, EventArgs e)
        {
           
            DeleteUnit();           
        }

        private void RMenu5_Click(object sender, EventArgs e)
        {
            EditClick();
        }

        private void RMenu4_Click(object sender, EventArgs e)
        {
            ViewClick();
        }

        private void RMenu3_Click(object sender, EventArgs e)
        {
            NewClick();

        }

   
        private bool CheckDuplicate(string code)
        {
            bool ck = false;

            //using (DataClasses1DataContext db = new DataClasses1DataContext())
            //{
            //    int i = (from ix in db.tb_Units where ix.UnitCode == code select ix).Count();
            //    if (i > 0)
            //        ck = false;
            //    else
            //        ck = true;
            //}
            return ck;
        }

        private bool AddUnit()
        {
            bool ck = false;


            return ck;
        }
        private bool DeleteUnit()
        {
            bool ck = false;
         


           

            return ck;
        }
        private void btnCancel_Click(object sender, EventArgs e)
        {
           
        }
        private void NewClick()
        {
            radGridView1.ReadOnly = false;
            radGridView1.AllowAddNewRow = false;
            //btnEdit.Enabled = false;
            btnView.Enabled = true;
            radGridView1.Rows.AddNew();
        }
        private void EditClick()
        {
            radGridView1.ReadOnly = false;
            //btnEdit.Enabled = false;
            btnView.Enabled = true;
            radGridView1.AllowAddNewRow = false;
        }
        private void ViewClick()
        {
           // radGridView1.ReadOnly = true;
            btnView.Enabled = false;
            //btnEdit.Enabled = true;
            radGridView1.AllowAddNewRow = false;
           
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            NewClick();
        }

        private void btnView_Click(object sender, EventArgs e)
        {
            //แสดงรายการ//
            //Open Page List Part//
            if (row >= 0)
            {
                WarehouseCheckPart whc = new WarehouseCheckPart(Convert.ToString(radGridView1.Rows[row].Cells["WONo"].Value),"");
                whc.Show();
            }
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {

            EditClick();
        }
        private void Saveclick()
        {
            if (MessageBox.Show("ต้องการบันทึก ?", "บันทึก", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                AddUnit();
              
            }
        }
        private void DeleteClick()
        {

        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            Saveclick();
        }


        private void radGridView1_CellEndEdit(object sender, Telerik.WinControls.UI.GridViewCellEventArgs e)
        {
            try
            {
         

                if(e.RowIndex>=0)
                {
                    if(e.ColumnIndex==radGridView1.Columns["Revision"].Index)
                    {
                        string RV=Convert.ToString(radGridView1.Rows[e.RowIndex].Cells["Revision"].Value);
                        string IV= Convert.ToString(radGridView1.Rows[e.RowIndex].Cells["InvoiceNo"].Value);
                        using (DataClasses1DataContext db = new DataClasses1DataContext())
                        {
                            tb_ExportList ep = db.tb_ExportLists.Where(s => s.InvoiceNo.Equals(IV)).FirstOrDefault();
                            if(ep!=null)
                            {
                                ep.Revision = RV;
                                db.SubmitChanges();
                            }
                        }
                    }
                }
        

            }
            catch(Exception ex) { }
        }

        private void Unit_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {


        }

        private void radGridView1_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {

            if (e.KeyData == (Keys.Control | Keys.S))
            {
                if (MessageBox.Show("ต้องการบันทึก ?", "บันทึก", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    AddUnit();
                   
                }
            }
            else if (e.KeyData == (Keys.Control | Keys.N))
            {
                if (MessageBox.Show("ต้องการสร้างใหม่ ?", "สร้างใหม่", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    NewClick();
                }
            }

        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            
                DeleteUnit();
                
            
        }

        int row = -1;
        private void radGridView1_CellClick(object sender, Telerik.WinControls.UI.GridViewCellEventArgs e)
        {
            row = e.RowIndex;
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            //dbClss.ExportGridCSV(radGridView1);
           dbClss.ExportGridXlSX(radGridView1);
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            OpenFileDialog op = new OpenFileDialog();
            op.Filter = "Spread Sheet files (*.csv)|*.csv|All files (*.csv)|*.csv";
            if (op.ShowDialog() == DialogResult.OK)
            {


                using (TextFieldParser parser = new TextFieldParser(op.FileName))
                {
                    dt.Rows.Clear();
                    parser.TextFieldType = FieldType.Delimited;
                    parser.SetDelimiters(",");
                    int a = 0;
                    int c = 0;
                    while (!parser.EndOfData)
                    {
                        //Processing row
                        a += 1;
                        DataRow rd = dt.NewRow();
                        // MessageBox.Show(a.ToString());
                        string[] fields = parser.ReadFields();
                        c = 0;
                        foreach (string field in fields)
                        {
                            c += 1;
                            //TODO: Process field
                            //MessageBox.Show(field);
                            if (a>1)
                            {
                                if(c==1)
                                    rd["UnitCode"] = Convert.ToString(field);
                                else if(c==2)
                                    rd["UnitDetail"] = Convert.ToString(field);
                                else if(c==3)
                                    rd["UnitActive"] = Convert.ToBoolean(field);

                            }
                            else
                            {
                                if (c == 1)
                                    rd["UnitCode"] = "";
                                else if (c == 2)
                                    rd["UnitDetail"] = "";
                                else if (c == 3)
                                    rd["UnitActive"] = false;




                            }

                            //
                            //rd[""] = "";
                            //rd[""]
                        }
                        dt.Rows.Add(rd);

                    }
                }
                if(dt.Rows.Count>0)
                {
                    dbClss.AddHistory(this.Name, "Import", "Import file CSV in to System", "");
                    ImportData();
                    MessageBox.Show("Import Completed.");

                    
                }
               
            }
        }

        private void ImportData()
        {
 
        }

        private void btnFilter1_Click(object sender, EventArgs e)
        {
            radGridView1.EnableFiltering = true;
        }

        private void btnUnfilter1_Click(object sender, EventArgs e)
        {
            radGridView1.EnableFiltering = false;
        }

        private void radMenuItem1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            DataLoad2();
        }

        private void ExportList_Load(object sender, EventArgs e)
        {
            // DateTime date = DateTime.Now;
            //var firstDayOfMonth = new DateTime(date.Year, date.Month, 1);
            //  var lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);
            defaultLoad();
            dtDate1.Value = DateTime.Today;
            dtDate2.Value = DateTime.Today;
            txtINv.Text = "";
            // radCheckBox1.Checked = false;
            DataLoad2();
        }
        private void defaultLoad()
        {
            using (DataClasses1DataContext db = new DataClasses1DataContext())
            {
             //   var listA=
                txtItemcat.DataSource = null;
                txtItemcat.DisplayMember = "itemCat";
                txtItemcat.ValueMember = "itemCat";
                txtItemcat.DataSource = db.sp_Z_3_pd_ListOnPC_itemCat().ToList();
                txtItemcat.Text = "";
            }
        }
        private void DataLoad2()
        {
            int ck = 0;
            using (DataClasses1DataContext db = new DataClasses1DataContext())
            {

                radGridView1.DataSource = db.sp_Z_3_pd_ListOnPC(radCheckBox1.Checked, dtDate1.Value, dtDate2.Value, txtINv.Text,txtItemcat.Text).ToList();
                foreach (var x in radGridView1.Rows)
                {

                    // x.Cells["dgvCodeTemp"].Value = x.Cells["UnitCode"].Value.ToString();
                    //  x.Cells["UnitCode"].ReadOnly = true;
                    //if (row >= 0 && row == ck && radGridView1.Rows.Count > 0)
                    //{

                    //    x.ViewInfo.CurrentRow = x;

                    //}
                    ck += 1;
                    x.Cells["No"].Value = ck;
                    
                }

            }
        }

        private void radGridView1_CellDoubleClick(object sender, Telerik.WinControls.UI.GridViewCellEventArgs e)
        {
            try
            {
                string Inv = radGridView1.Rows[e.RowIndex].Cells["InvoiceNo"].Value.ToString();
                ExShipment ex = new ExShipment(Inv);
                ex.Show();
            }
            catch { }
        }

        private void radButtonElement1_Click(object sender, EventArgs e)
        {
            InvoiceEx_Master iv = new InvoiceEx_Master();
            iv.Show();
        }

        private void radButtonElement2_Click(object sender, EventArgs e)
        {
            try
            {
                InvoiceEx_MasterMap iv = new InvoiceEx_MasterMap();
                iv.Show();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void radButtonElement3_Click(object sender, EventArgs e)
        {
            ExportManual exm = new ExportManual();
            exm.ShowDialog();
        }

        private void radButtonElement4_Click(object sender, EventArgs e)
        {
            try
            {
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    db.sp_80_Prod_Delete();
                    string QBarCode;
                    foreach(var x in radGridView1.Rows)
                    {
                        QBarCode = "";
                        QBarCode = Convert.ToString(x.Cells["WONo"].Value).ToString();
                      //  MessageBox.Show(QBarCode);
                        tb_ProductionTempPrint pdi = new tb_ProductionTempPrint();
                        pdi.WONo = x.Cells["WONo"].Value.ToString();
                        pdi.PartNo = x.Cells["PartNo"].Value.ToString();
                        pdi.WODate=Convert.ToDateTime(x.Cells["WODate"].Value);
                        pdi.WOQty = Convert.ToDecimal(x.Cells["Quantity"].Value);
                        pdi.LOTNo = x.Cells["LotNo"].Value.ToString();
                        pdi.ItemCat = x.Cells["ItemCat"].Value.ToString();
                        byte[] barcode = dbClss.SaveQRCode2D(QBarCode);
                        pdi.Barcode = barcode;
                        pdi.DN = x.Cells["DN"].Value.ToString();
                        pdi.Remark = x.Cells["Remark"].Value.ToString();
                        pdi.Status = x.Cells["CheckStatus"].Value.ToString();
                        pdi.Description = x.Cells["Description"].Value.ToString();
                        pdi.Seq = Convert.ToInt32(x.Cells["Seq"].Value.ToString());
                        
                        db.tb_ProductionTempPrints.InsertOnSubmit(pdi);
                        db.SubmitChanges();
                    }
                    Report.Reportx1.WReport = "QCQRCode";
                    Report.Reportx1.Value = new string[1];
                    Report.Reportx1.Value[0] = "";//Datex
                    Report.Reportx1 op = new Report.Reportx1("WH_ReportCheck.rpt");
                    op.Show();
                }
            }
            catch { }
        }

        private void radButtonElement5_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("ต้องการอัพเดตหรือไม่ ?", "อัพเดต", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    using (DataClasses1DataContext db = new DataClasses1DataContext())
                    {
                        foreach (var x in radGridView1.Rows)
                        {
                            db.sp_Z_3_pd_List_UPdate(x.Cells["WONo"].Value.ToString());
                        }
                    }
                    MessageBox.Show("Update Completed");
                    DataLoad2();
                }
            }
            catch { }
        }
        int CCRow = 0;
        private void radGridView1_RowFormatting(object sender, Telerik.WinControls.UI.RowFormattingEventArgs e)
        {
            try
            {
                if (CCRow > 5000)
                {
                    CCRow += 1;
                }
                else
                {
                    e.RowElement.ResetValue(LightVisualElement.BackColorProperty, ValueResetFlags.Local);
                    e.RowElement.ResetValue(LightVisualElement.GradientStyleProperty, ValueResetFlags.Local);
                    e.RowElement.ResetValue(LightVisualElement.DrawFillProperty, ValueResetFlags.Local);

                    if (Convert.ToString(e.RowElement.RowInfo.Cells["CheckStatus"].Value).Equals("Waiting"))
                    {

                        e.RowElement.ResetValue(LightVisualElement.BackColorProperty, ValueResetFlags.Local);
                        e.RowElement.ResetValue(LightVisualElement.GradientStyleProperty, ValueResetFlags.Local);
                        e.RowElement.ResetValue(LightVisualElement.DrawFillProperty, ValueResetFlags.Local);

                    }
                    else if (Convert.ToString(e.RowElement.RowInfo.Cells["CheckStatus"].Value).Equals("Partial")
                        )
                    {

                        e.RowElement.DrawFill = true;
                        e.RowElement.GradientStyle = GradientStyles.Solid;
                        e.RowElement.BackColor = Color.Yellow;

                    }
                    else if (Convert.ToString(e.RowElement.RowInfo.Cells["CheckStatus"].Value).Equals("Completed")                      
                        )
                    {

                        e.RowElement.DrawFill = true;
                        e.RowElement.GradientStyle = GradientStyles.Solid;
                        e.RowElement.BackColor = Color.LightGreen;

                    }            
                    else
                    {
                        e.RowElement.ResetValue(LightVisualElement.BackColorProperty, ValueResetFlags.Local);
                        e.RowElement.ResetValue(LightVisualElement.GradientStyleProperty, ValueResetFlags.Local);
                        e.RowElement.ResetValue(LightVisualElement.DrawFillProperty, ValueResetFlags.Local);
                    }
                }
            }
            catch { }
        }

        private void radButtonElement6_Click(object sender, EventArgs e)
        {
            try
            {
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    db.sp_Z_3_pd_ListGraph(dbClss.UserID, dtDate1.Value, dtDate2.Value);

                    Report.Reportx1.WReport = "WHPDA";
                    Report.Reportx1.Value = new string[1];
                    Report.Reportx1.Value[0] = dbClss.UserID;
                    Report.Reportx1 op = new Report.Reportx1("WH_ReportGPH.rpt");
                    op.Show();

                }
            }
            catch { }
        }
    }
}
