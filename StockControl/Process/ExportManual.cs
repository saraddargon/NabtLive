using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Linq;
using Microsoft.VisualBasic.FileIO;
using Telerik.WinControls.UI;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
namespace StockControl
{
    public partial class ExportManual : Telerik.WinControls.UI.RadRibbonForm
    {
        public ExportManual()
        {
            InitializeComponent();
        }
        public ExportManual(int idx,string Editx,string ListNoPrevx)
        {
            InitializeComponent();
            id = idx;
            Edit = Editx;
            ListNoPrev = ListNoPrevx;
            
        }
        string Edit = "";
        string ListNoPrev = "";
        int id = 0;
        decimal ListOld = 0;
        //private int RowView = 50;
        //private int ColView = 10;
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
            //dt.Columns.Add(new DataColumn("UnitCode", typeof(string)));
            //dt.Columns.Add(new DataColumn("UnitDetail", typeof(string)));
            //dt.Columns.Add(new DataColumn("UnitActive", typeof(bool)));
        }
        private void Clear()
        {
            txtPreListNo.ReadOnly = false;
            txtInvoiceNo.Text = "";
            txtGroupP.Text = "";
            txtPreListNo.Text = "None";
            txtPreListNo.ReadOnly = true;
            dtShippingDate.Value = DateTime.Now;
            txtPartCode.Text = "";
            txtPartName.Text = "";
            txtQty.Text = "";
            txtQtyofPL.Text = "";
            txtQtyofTAG.Text = "";
            txtLotNo.Text = "";
            txtSTDPacking.Text = "";
            txtStatus.Text = "";
            txtOrderNo.Text = "";
            txtCustomer.Text = "";
            txtPallet.Text = "";
            txtOldQty.Text = "";
        }
        private void Unit_Load(object sender, EventArgs e)
        {
            setDT();
           // DefaultLoad();
        }
        private void DefaultLoad()
        {
            try
            {
                if (!id.Equals(0))
                {

                    using (DataClasses1DataContext db = new DataClasses1DataContext())
                    {
                        if (Edit.Equals("Add"))
                        {
                            lblStatus.Text = "Add";
                            tb_ExportDetail dd = db.tb_ExportDetails.Where(d => d.id == id).FirstOrDefault();
                            if (dd != null)
                            {
                                txtInvoiceNo.Text = dd.InvoiceNo;
                                txtGroupP.Text = dd.GroupP.ToString();
                                txtPreListNo.Text = Convert.ToInt32(dd.ListNo).ToString();
                                txtListNo.Text = Convert.ToDecimal(dd.ListNo + 1).ToString();
                                txtPreListNo.ReadOnly = false;
                                dtShippingDate.Value = DateTime.Now;
                            }

                        }
                        else
                        {
                            lblStatus.Text = "Edit";
                            tb_ExportDetail dd = db.tb_ExportDetails.Where(d => d.id == id).FirstOrDefault();
                            if (dd != null)
                            {
                                ListOld = Convert.ToDecimal(dd.ListNo);
                                txtInvoiceNo.Text = dd.InvoiceNo;
                                txtGroupP.Text = dd.GroupP.ToString();
                                txtPreListNo.Text = "None";
                                txtPreListNo.ReadOnly = true;
                                txtListNo.Text = Convert.ToInt32(dd.ListNo).ToString();
                                txtPartCode.Text = dd.PartNo.ToString();
                                txtPartName.Text = dd.PartName.ToString();
                                txtQty.Text = dd.Qty.ToString();
                                txtQtyofPL.Text = dd.QtyOfPL.ToString();
                                txtQtyofTAG.Text = dd.QtyOfTAG.ToString();
                                txtLotNo.Text = dd.LotNo.ToString();
                                txtSTDPacking.Text = dd.STDPacking.ToString();
                                txtStatus.Text = dd.Status.ToString();
                                txtOrderNo.Text = dd.OrderNo.ToString();
                                txtCustomer.Text = dd.Customer.ToString();
                                txtPallet.Text = dd.PalletNo.ToString();
                                txtOldQty.Text = Convert.ToInt32(dd.OldQty).ToString();
                                dtShippingDate.Value = Convert.ToDateTime(dd.ShippingDate);


                            }
                        }
                    }
                    LoadData();
                    txtListNo.Focus();
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }
        private void LoadData()
        {
            this.Cursor = Cursors.WaitCursor;
            try
            {
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    tb_ExportList ex = db.tb_ExportLists.Where(es => es.InvoiceNo == txtInvoiceNo.Text).FirstOrDefault();
                    if (ex != null)
                    {
                       // txtLoadDate.Text = Convert.ToDateTime(ex.LoadDate).ToString("dd/MM/yyyy");


                        //LoadDetail//
                        //radGridView1.DataSource = null;
                        //int Group = 0;
                        //int.TryParse(txtGroupP.Text, out Group);

                        //var pl = (from ix in db.sp_013_selectExportList_DetailGroup(txtInvoiceNo.Text,Group) select ix).ToList();
                        //if (pl.Count > 0)
                        //{
                        //    //txtTotalPallet.Text = pl.FirstOrDefault().TotalPallet.ToString();
                        //    radGridView1.DataSource = pl;
                        //    int[] ts = new int[] { 5 };
                        //    MergeVertically(radGridView1, ts);

                        //}

                        //int ck = 0;
                        // foreach()
                        // MergeVertical2();
                    }
                }
            }
            catch (Exception ex) { this.Cursor = Cursors.Default; MessageBox.Show(ex.Message); }
            this.Cursor = Cursors.Default;
        }
        private void setCellBorders(GridViewCellInfo cell, Color color)
        {
            cell.Style.CustomizeBorder = true;
            cell.Style.BorderBoxStyle = Telerik.WinControls.BorderBoxStyle.FourBorders;
            cell.Style.BorderLeftColor = color;
            cell.Style.BorderRightColor = color;
            cell.Style.BorderBottomColor = color;
            if (cell.Style.BorderTopColor != Color.Transparent)
            {
                cell.Style.BorderTopColor = color;
            }
        }

        private void MergeVertically(RadGridView radGridView, int[] columnIndexes)
        {
            GridViewRowInfo Prev = null;
            foreach (GridViewRowInfo item in radGridView.Rows)
            {
                if (Prev != null)
                {
                    string firstCellText = string.Empty;
                    string secondCellText = string.Empty;

                    foreach (int i in columnIndexes)
                    {
                        GridViewCellInfo firstCell = Prev.Cells[i];
                        GridViewCellInfo secondCell = item.Cells[i];

                        firstCellText = (firstCell != null && firstCell.Value != null ? firstCell.Value.ToString() : string.Empty);
                        secondCellText = (secondCell != null && secondCell.Value != null ? secondCell.Value.ToString() : string.Empty);

                        setCellBorders(firstCell, Color.FromArgb(209, 225, 245));
                        setCellBorders(secondCell, Color.FromArgb(209, 225, 245));
                        if (secondCellText.Equals(0) || secondCellText.Equals("0") || secondCellText.Equals(""))
                        {
                            firstCell.Style.BorderBottomColor = Color.Transparent;
                            secondCell.Style.BorderTopColor = Color.Transparent;
                            secondCell.Style.ForeColor = Color.Transparent;
                            Prev = item;
                        }
                        else
                        {
                            secondCell.Style.ForeColor = Color.Black;
                            Prev = item;
                            break;
                        }

                        //if (firstCellText == secondCellText)
                        //{
                        //    firstCell.Style.BorderBottomColor = Color.Transparent;
                        //    secondCell.Style.BorderTopColor = Color.Transparent;
                        //    secondCell.Style.ForeColor = Color.Transparent;
                        //    Prev = item;
                        //}
                        //else
                        //{
                        //    secondCell.Style.ForeColor = Color.Black;
                        //    Prev = item;
                        //    break;
                        //}
                    }
                }
                else
                {
                    Prev = item;
                }
            }
        }

        private void RMenu6_Click(object sender, EventArgs e)
        {
           
           // DeleteUnit();
            //DataLoad();
        }

        private void RMenu5_Click(object sender, EventArgs e)
        {
            //EditClick();
        }

        private void RMenu4_Click(object sender, EventArgs e)
        {
           // ViewClick();
        }

        private void RMenu3_Click(object sender, EventArgs e)
        {
           // NewClick();

        }

        private void DataLoad()
        {
           
            
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
            //int C = 0;
            //try
            //{


            //    radGridView1.EndEdit();
            //    using (DataClasses1DataContext db = new DataClasses1DataContext())
            //    {
            //        foreach (var g in radGridView1.Rows)
            //        {
            //            if (!Convert.ToString(g.Cells["UnitCode"].Value).Equals(""))
            //            {
            //                if (Convert.ToString(g.Cells["dgvC"].Value).Equals("T"))
            //                {
                               
            //                    if (Convert.ToString(g.Cells["dgvCodeTemp"].Value).Equals(""))
            //                    {
            //                       // MessageBox.Show("11");
                                    
            //                        tb_Unit u = new tb_Unit();
            //                        u.UnitCode = Convert.ToString(g.Cells["UnitCode"].Value);
            //                        u.UnitActive = Convert.ToBoolean(g.Cells["UnitActive"].Value);
            //                        u.UnitDetail= Convert.ToString(g.Cells["UnitDetail"].Value);
            //                        db.tb_Units.InsertOnSubmit(u);
            //                        db.SubmitChanges();
            //                        C += 1;
            //                        dbClss.AddHistory(this.Name, "เพิ่ม", "Insert Unit Code [" + u.UnitCode+"]","");
            //                    }
            //                    else
            //                    {
                                   
            //                        var unit1 = (from ix in db.tb_Units
            //                                     where ix.UnitCode == Convert.ToString(g.Cells["dgvCodeTemp"].Value)
            //                                     select ix).First();
            //                           unit1.UnitDetail = Convert.ToString(g.Cells["UnitDetail"].Value);
            //                           unit1.UnitActive = Convert.ToBoolean(g.Cells["UnitActive"].Value);
                                    
            //                        C += 1;

            //                        db.SubmitChanges();
            //                        dbClss.AddHistory(this.Name, "แก้ไข", "Update Unit Code [" + unit1.UnitCode+"]","");

            //                    }
            //                }
            //            }
            //        }
            //    }
            //}
            //catch (Exception ex) { MessageBox.Show(ex.Message);
            //    dbClss.AddError("AddUnit", ex.Message, this.Name);
            //}

            //if (C > 0)
            //    MessageBox.Show("บันทึกสำเร็จ!");

            return ck;
        }
        private bool DeleteUnit()
        {
            bool ck = false;
         
            //int C = 0;
            //try
            //{
                
            //    if (row >= 0)
            //    {
            //        string CodeDelete = Convert.ToString(radGridView1.Rows[row].Cells["UnitCode"].Value);
            //        string CodeTemp = Convert.ToString(radGridView1.Rows[row].Cells["dgvCodeTemp"].Value);
            //        radGridView1.EndEdit();
            //        if (MessageBox.Show("ต้องการลบรายการ ( "+ CodeDelete+" ) หรือไม่ ?", "ลบรายการ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            //        {
            //            using (DataClasses1DataContext db = new DataClasses1DataContext())
            //            {

            //                if (!CodeDelete.Equals(""))
            //                {
            //                    if (!CodeTemp.Equals(""))
            //                    {

            //                        var unit1 = (from ix in db.tb_Units
            //                                     where ix.UnitCode == CodeDelete
            //                                     select ix).ToList();
            //                        foreach (var d in unit1)
            //                        {
            //                            db.tb_Units.DeleteOnSubmit(d);
            //                            dbClss.AddHistory(this.Name, "ลบ Unit", "Delete Unit Code ["+d.UnitCode+"]","");
            //                        }
            //                        C += 1;



            //                        db.SubmitChanges();
            //                    }
            //                }

            //            }
            //        }
            //    }
            //}

            //catch (Exception ex) { MessageBox.Show(ex.Message);
            //    dbClss.AddError("DeleteUnit", ex.Message, this.Name);
            //}

            //if (C > 0)
            //{
            //        row = row - 1;
            //        MessageBox.Show("ลบรายการ สำเร็จ!");
            //}
              

           

            return ck;
        }
        private void btnCancel_Click(object sender, EventArgs e)
        {
            //DataLoad();
            radLabel24.Text = "0";
            Clear();
            //LoadData();
        }
        private void clear()
        {

        }
        private void NewClick()
        {
            //radGridView1.ReadOnly = false;
            //radGridView1.AllowAddNewRow = false;
            //btnEdit.Enabled = false;
           // btnView.Enabled = true;
           // radGridView1.Rows.AddNew();
        }
        private void EditClick()
        {
          //  radGridView1.ReadOnly = false;
            //btnEdit.Enabled = false;
            //btnView.Enabled = true;
           // radGridView1.AllowAddNewRow = false;
        }
        private void ViewClick()
        {
           // radGridView1.ReadOnly = true;
           // btnView.Enabled = false;
            //btnEdit.Enabled = true;
           // radGridView1.AllowAddNewRow = false;
            DataLoad();
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            //NewClick();
        }

        private void btnView_Click(object sender, EventArgs e)
        {
            //ViewClick();
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {

            //EditClick();
        }
        private void Saveclick()
        {

            
            if (MessageBox.Show("ต้องการบันทึก ?", "บันทึก", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    this.Cursor = Cursors.WaitCursor;
                    int Qty = 0;
                    int QtyofPL = 0;
                    int OldQty = 0;
                    decimal QtyofTAG = 0;
                    int STDPacking = 0;
                    int GroupP = 0;
                    decimal listNo = 0;
                    int preList = 0;
                    int.TryParse(txtQty.Text, out Qty);
                    int.TryParse(txtQtyofPL.Text, out QtyofPL);
                    int.TryParse(txtGroupP.Text, out GroupP);
                    int.TryParse(txtSTDPacking.Text, out STDPacking);
                    decimal.TryParse(txtQtyofTAG.Text, out QtyofTAG);
                    decimal.TryParse(txtListNo.Text, out listNo);
                    int.TryParse(txtPreListNo.Text, out preList);
                    int.TryParse(txtOldQty.Text, out OldQty);

                    using (DataClasses1DataContext db = new DataClasses1DataContext())
                    {
                        if (Edit.Equals("Add"))
                        {
                            tb_ExportDetail ul = db.tb_ExportDetails.Where(u => u.InvoiceNo == txtInvoiceNo.Text && u.ListNo == listNo).FirstOrDefault();
                            if (ul != null)
                            {
                                ul.ListNo = Convert.ToDecimal(ul.ListNo) + Convert.ToDecimal(0.1);
                                db.SubmitChanges();

                            }
                            tb_ExportDetail ud = new tb_ExportDetail();
                            ud.InvoiceNo = txtInvoiceNo.Text;
                            ud.PartNo = txtPartCode.Text;
                            ud.PartName = txtPartName.Text;
                            ud.Qty = Qty;
                            ud.QtyOfPL = QtyofPL;
                            ud.QtyOfTAG = QtyofTAG;
                            ud.STDPacking = STDPacking;
                            ud.Status = txtStatus.Text;
                            ud.Customer = txtCustomer.Text;
                            ud.ListNo = listNo;
                            ud.LotNo = txtLotNo.Text;
                            ud.GroupP = GroupP.ToString();
                            ud.OrderNo = txtOrderNo.Text;
                            ud.Status2 = "W";
                            ud.PrintType = "";
                          //  ud.PalletNo = "";
                            ud.SS = 1;
                            ud.ConfirmFlag = false;
                            ud.ShipFlag = false;
                            ud.SNP = STDPacking;
                            ud.ShippingDate = dtShippingDate.Value;
                            ud.PalletNo = txtPallet.Text;
                            ud.OldQty = Qty;
                            db.tb_ExportDetails.InsertOnSubmit(ud);
                            db.SubmitChanges();


                        }
                        else
                        {
                            tb_ExportDetail ed = db.tb_ExportDetails.Where(ee => ee.id == id && ee.SS != 3).FirstOrDefault();
                            if (ed != null)
                            {

                                ///Update ListNo//
                                if (ListOld != listNo)
                                {
                                    tb_ExportDetail ul = db.tb_ExportDetails.Where(u => u.InvoiceNo == txtInvoiceNo.Text && u.ListNo == listNo).FirstOrDefault();
                                    if(ul!=null)
                                    {
                                        ul.ListNo = Convert.ToDecimal(ul.ListNo)+Convert.ToDecimal(0.1);
                                        db.SubmitChanges();

                                    }

                                }

                                ed.Qty = Qty;
                                ed.QtyOfPL = QtyofPL;
                                ed.QtyOfTAG = QtyofTAG;
                                ed.STDPacking = STDPacking;
                                ed.Status = txtStatus.Text;
                                ed.Customer = txtCustomer.Text;
                                ed.ListNo = listNo;
                                ed.LotNo = txtLotNo.Text;
                                ed.GroupP = GroupP.ToString();
                                ed.OrderNo = txtOrderNo.Text;
                                ed.ShippingDate = dtShippingDate.Value;
                                ed.PartNo = txtPartCode.Text;
                                ed.PartName = txtPartName.Text;
                                ed.PalletNo = txtPallet.Text;
                                ed.OldQty = OldQty;
                                db.SubmitChanges();
                                if (ListOld != listNo)
                                {
                                    UpdateListNo(ListOld,listNo);
                                    ListOld = listNo;
                                }

                            }

                        }


                        dbClss.AddHistory("ExportEditAdd", lblStatus.Text, "เพิ่ม/แก้ไข " + txtInvoiceNo.Text + " เข้าระบบ", "List No:" + txtListNo.Text.ToString());
                        MessageBox.Show("Save Completed!");
                    }
                }
                catch (Exception ex)
                {
                    this.Cursor = Cursors.Default;
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            this.Cursor = Cursors.Default;

        }
        private void UpdateListNo(decimal ListOld, decimal ListNew)
        {
            try
            {
                decimal CountRow = 0;
                int Fix = 0;
                int AL = 0;
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    var getList = db.tb_ExportDetails.Where(ee => ee.id == id).OrderBy(o=>o.ListNo).OrderBy(o=>o.id).ToList();
                    foreach(var rd in getList)
                    {
                        CountRow += 1;
                        if (ListOld==CountRow || (Convert.ToDecimal(rd.ListNo)%1)!=0)
                        {
                            Fix = 1;                            
                        }

                        if(Fix==1 && rd.id!=id)
                        {
                            tb_ExportDetail ed = db.tb_ExportDetails.Where(es => es.id == rd.id).FirstOrDefault();
                            if(ed!=null)
                            {
                                ed.ListNo = CountRow;
                                db.SubmitChanges();
                            }

                        }
                    }
                }
            }
            catch { }
        }
        private void DeleteClick()
        {

        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            Saveclick();
           // btnSave.Enabled = false;
        }


        private void radGridView1_CellEndEdit(object sender, Telerik.WinControls.UI.GridViewCellEventArgs e)
        {

        }

        private void Unit_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {


        }

        private void radGridView1_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {

        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            
                //DeleteUnit();
                //DataLoad();
            
        }

        int row = -1;
        private void radGridView1_CellClick(object sender, Telerik.WinControls.UI.GridViewCellEventArgs e)
        {
            row = e.RowIndex;
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            //dbClss.ExportGridCSV(radGridView1);
          // dbClss.ExportGridXlSX(radGridView1);
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

                    DataLoad();
                }
               
            }
        }

        private void ImportData()
        {
            //try
            //{
            //    using (DataClasses1DataContext db = new DataClasses1DataContext())
            //    {
                   
            //        foreach (DataRow rd in dt.Rows)
            //        {
            //            if (!rd["UnitCode"].ToString().Equals(""))
            //            {

            //                var x = (from ix in db.tb_Units where ix.UnitCode.ToLower().Trim() == rd["UnitCode"].ToString().ToLower().Trim() select ix).FirstOrDefault();

            //                if(x==null)
            //                {
            //                    tb_Unit ts = new tb_Unit();
            //                    ts.UnitCode = Convert.ToString(rd["UnitCode"].ToString());
            //                    ts.UnitDetail = Convert.ToString(rd["UnitDetail"].ToString());
            //                    ts.UnitActive = Convert.ToBoolean(rd["UnitActive"].ToString());
            //                    db.tb_Units.InsertOnSubmit(ts);
            //                    db.SubmitChanges();
            //                }
            //                else
            //                {
            //                    x.UnitDetail = Convert.ToString(rd["UnitDetail"].ToString());
            //                    x.UnitActive = Convert.ToBoolean(rd["UnitActive"].ToString());
            //                    db.SubmitChanges();

            //                }

                       
            //            }
            //        }
                   
            //    }
            //}
            //catch(Exception ex) { MessageBox.Show(ex.Message);
            //    dbClss.AddError("InportData", ex.Message, this.Name);
            //}
        }

        private void btnFilter1_Click(object sender, EventArgs e)
        {
            //radGridView1.EnableFiltering = true;
        }

        private void btnUnfilter1_Click(object sender, EventArgs e)
        {
           // radGridView1.EnableFiltering = false;
        }

        private void radMenuItem1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void radTextBox12_KeyPress(object sender, KeyPressEventArgs e)
        {
           // dbClss.checkDigit(e);
        }

        private void radTextBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
           // dbClss.checkDigit(e);
        }

        private void radTextBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
           // dbClss.checkDigit(e);
        }

        private void radTextBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
           // dbClss.checkDigit(e);
        }

        private void radTextBox8_KeyPress(object sender, KeyPressEventArgs e)
        {
            //dbClss.CheckDigitDecimal(e);
        }

        private void radTextBox11_KeyPress(object sender, KeyPressEventArgs e)
        {
            //dbClss.checkDigit(e);
        }

        private void radButton1_Click(object sender, EventArgs e)
        {
            LoadData();
        }

        private void txtPartCode_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                LoadDataPart();
            }
        }

        private void LoadDataPart()
        {
            if (!txtOrderNo.Text.Equals("") && !txtPartCode.Text.Equals(""))
            {
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    var splist1 = db.sp_035_LoadSaleOrder_Dynamics(txtOrderNo.Text, txtPartCode.Text).FirstOrDefault();
                    if (splist1 != null)
                    {
                        txtPartName.Text = splist1.NAME;
                        txtCustomer.Text = splist1.CUSSHORTNAME.ToString();
                        txtSTDPacking.Text = Convert.ToInt32(splist1.SNP).ToString();
                        txtStatus.Text = "";
                        txtQty.Text = Convert.ToInt32(splist1.KVOL).ToString();
                        txtLotNo.Text = "";
                        txtQtyofPL.Text = "0";
                        txtQtyofTAG.Text = "1";
                        dtShippingDate.Value = Convert.ToDateTime(splist1.ShippingDate);
                       
                    }
                }
            }
            else
            {
                MessageBox.Show("[Sale Order No.] or [Part No.] is Empty!");
            }
        }

        private void radButtonElement1_Click(object sender, EventArgs e)
        {
            try
            {
                string QRCode = "";
                using (DataClasses1DataContext db= new DataClasses1DataContext())
                {
                    txtPartName.Text = db.getItemNo(txtPartName.Text);
                    db.sp_014_DeletePrintTAG(txtInvoiceNo.Text);
                    QRCode = "";
                    QRCode = txtOrderNo.Text + "," + txtPallet.Text + "," + txtInvoiceNo.Text + ",";
                    QRCode = QRCode + txtPartCode.Text + "," + txtQty.Text + "," +"1of" + txtQtyofPL.Text + "," +txtLotNo.Text;
                    byte[] barcode = dbClss.SaveQRCode2D(QRCode);

                    tb_ExportPrintTAG ep = new tb_ExportPrintTAG();
                    ep.CustomerAddress = "5-1 Kanaya, Murayama, Yamagata, 995-0004 Japan";// ed.CustomerAddress.ToString();
                                                                                          //MessageBox.Show("OK");
                    ep.CustomerItemName = ""; //// ed.CustItem.ToString();                               
                    ep.CustomerItemNo = Convert.ToString(db.getItemCSTM_Dynamics(txtPartCode.Text, ""));
                    ep.CustomerName = "Nabtesco Autmotive Corporation";// Convert.ToString(db.getItemCSTMName(ed.Customer));

                    if (txtOldQty.Text.ToUpper().Equals("INDIA"))
                    {
                        ep.CustomerName = "MINDA NABTESCO AUTOMOTIVE PVT LTD";
                        ep.CustomerAddress = "Plot no-191 sector-8 IMT  Manesar ,distt- Gurgaon- 122050 State Haryana.";
                    }
                    ep.InvoiceNo = txtInvoiceNo.Text;


                    ep.LOTNo = txtLotNo.Text;
                    ep.QRCode = barcode;

                    ep.Qty = Convert.ToDecimal(txtQty.Text);
                    if (txtGroupP.Text.Equals(""))
                    {
                        txtGroupP.Text = "1";
                    }
                    ep.GroupP = txtGroupP.Text;
                    ep.ShippingDate = dtShippingDate.Value;
                    ep.TotalPLOf = Convert.ToInt32(txtPallet.Text);
                    ep.TotalPLOfQty = Convert.ToInt32(txtQtyofPL.Text);
                    ep.PLOfQty = Convert.ToInt32(txtQtyofPL.Text);
                    ep.PLOf = Convert.ToInt32(txtQtyofPL.Text);
                    ep.PartCode = txtPartCode.Text;
                    ep.PartName = txtPartName.Text;
                    ep.Country = txtOldQty.Text;
                    ep.CountrySize = txtSTDPacking.Text;


                    db.tb_ExportPrintTAGs.InsertOnSubmit(ep);
                    db.SubmitChanges();

                }
                Report.Reportx1.WReport = "PrintEXTAG";
                Report.Reportx1.Value = new string[1];
                Report.Reportx1.Value[0] = txtInvoiceNo.Text;
                Report.Reportx1 op = new Report.Reportx1("ExInvoiceTAX.rpt");
            }
            catch { }
        }
        DataTable dt_d = new DataTable();
        private void setDT()
        {
            dt_d = new DataTable();
            dt_d.Columns.Add(new DataColumn("PartNo", typeof(string)));
            dt_d.Columns.Add(new DataColumn("CustItem", typeof(string)));
            dt_d.Columns.Add(new DataColumn("Qty", typeof(decimal)));
            dt_d.Columns.Add(new DataColumn("LotNo", typeof(string)));
        }
        private void btnImport_Click_1(object sender, EventArgs e)
        {


            try
            {
                this.Cursor = Cursors.WaitCursor;
                string PartFile = "";
                dt_d.Rows.Clear();
                openFileDialog1.FilterIndex = 2;
                openFileDialog1.RestoreDirectory = true;

                openFileDialog1.DefaultExt = "*.xls";
                openFileDialog1.AddExtension = true;
                openFileDialog1.FileName = "";
                openFileDialog1.Filter = "Excel 2003-2010  (*.xls,*.xlsx,*.csv)|*.xls;*xlsx;*.csv";

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    PartFile = openFileDialog1.FileName;
                    //string name = Path.GetFileName(openFileDialog1.FileName);
                    string Exten = Path.GetExtension(openFileDialog1.FileName);
                    if (Exten.ToUpper() == ".XLS" || Exten.ToUpper() == ".XLSX")
                        Import_Excel(openFileDialog1.FileName);
                    //else if (Exten.ToUpper() == ".CSV")
                    //    Import_CSV(openFileDialog1.FileName);

                    if (dt_d.Rows.Count > 0)
                        radLabel24.Text = dt_d.Rows.Count.ToString();
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

                
                int icount = 0;

                //int Sheet4 = 0;
                for (int j = 2; j < 1003; j++)
                {
                

                    System.Array myvalues;
                    Excel.Range range = worksheet.get_Range("A" + j.ToString(), "D" + j.ToString());
                    myvalues = (System.Array)range.Cells.Value;
                    string[] strArray = ConvertToStringArray(myvalues);
                    if (!Convert.ToString(strArray[0]).Equals("")
                        //!Convert.ToString(strArray[2]).Equals("")
                        )
                    {
                        GetDataSystem2(Convert.ToString(strArray[0]).Trim() //PartNo
                            , Convert.ToString(strArray[1]).Trim()//CustItemName
                            , Convert.ToDecimal(strArray[2])//Qty
                            , Convert.ToString(strArray[3]).Trim()//Lot                           
                            );
                    }
                    else
                        break;
                }
              

                //excelBook.Save();
                //excelApp.Quit();
                releaseObject(worksheet);
                releaseObject(excelApp);
                //Marshal.FinalReleaseComObject(worksheet);


            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }

        }
        private void GetDataSystem2(string PartNapt, string PartCust, decimal Qty, string LotNo)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;

               int RowIndex_temp = (dt_d.Rows.Count);
                DataRow dr = dt_d.NewRow();

                string Status = "OK";

                dr["PartNo"] = PartNapt;

                //if (dbClss.TSt(CustItemNo).Trim() != "" && Status == "OK")
                dr["CustItem"] = PartCust;
                //else
                //    Status = "NG";
                //if (dbClss.TSt(CustItemName).Trim() != "" && Status == "OK")
                dr["Qty"] = Qty;
                //else
                //    Status = "NG";

                dr["LotNo"] = LotNo;
                if (Status == "OK")
                {
                    dt_d.Rows.Add(dr);
                    RowIndex_temp += 1;
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            //finally { this.Cursor = Cursors.Default; }
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

        private void radButtonElement2_Click(object sender, EventArgs e)
        {
            try
            {
                string QRCode = "";
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    txtPartName.Text = db.getItemNo(txtPartName.Text);
                    db.sp_014_DeletePrintTAG(txtInvoiceNo.Text);
                    foreach (DataRow rd in dt_d.Rows)
                    {
                        QRCode = "";
                        QRCode = txtOrderNo.Text + "," + txtPallet.Text + "," + txtInvoiceNo.Text + ",";
                        QRCode = QRCode + rd["PartNo"].ToString() + "," + Convert.ToInt32(rd["Qty"].ToString()) + "," + "1of" + txtQtyofPL.Text + "," + rd["LotNo"].ToString();
                        byte[] barcode = dbClss.SaveQRCode2D(QRCode);

                        tb_ExportPrintTAG ep = new tb_ExportPrintTAG();
                        ep.CustomerAddress = "5-1 Kanaya, Murayama, Yamagata, 995-0004 Japan";// ed.CustomerAddress.ToString();
                                                                                              //MessageBox.Show("OK");
                        ep.CustomerItemName = ""; //// ed.CustItem.ToString();                               
                        ep.CustomerItemNo = Convert.ToString(db.getItemCSTM_Dynamics(rd["PartNo"].ToString(), ""));
                        ep.CustomerName = "Nabtesco Autmotive Corporation";// Convert.ToString(db.getItemCSTMName(ed.Customer));

                        if (txtOldQty.Text.ToUpper().Equals("INDIA"))
                        {
                            ep.CustomerName = "MINDA NABTESCO AUTOMOTIVE PVT LTD";
                            ep.CustomerAddress = "Plot no-191 sector-8 IMT  Manesar ,distt- Gurgaon- 122050 State Haryana.";
                        }
                        ep.InvoiceNo = txtInvoiceNo.Text;


                        ep.LOTNo = rd["LotNo"].ToString();
                        ep.QRCode = barcode;

                        ep.Qty = Convert.ToDecimal(rd["Qty"].ToString());
                        if (txtGroupP.Text.Equals(""))
                        {
                            txtGroupP.Text = "1";
                        }
                        ep.GroupP = txtGroupP.Text;
                        ep.ShippingDate = dtShippingDate.Value;
                        ep.TotalPLOf = Convert.ToInt32(txtPallet.Text);
                        ep.TotalPLOfQty = Convert.ToInt32(txtQtyofPL.Text);
                        ep.PLOfQty = Convert.ToInt32(txtQtyofPL.Text);
                        ep.PLOf = Convert.ToInt32(txtQtyofPL.Text);
                        ep.PartCode = rd["PartNo"].ToString();
                        ep.PartName = db.getItemNo(rd["PartNo"].ToString());
                        ep.Country = txtOldQty.Text;
                        ep.CountrySize = txtSTDPacking.Text;


                        db.tb_ExportPrintTAGs.InsertOnSubmit(ep);
                        db.SubmitChanges();
                    }
                }

                Report.Reportx1.WReport = "PrintEXTAG";
                Report.Reportx1.Value = new string[1];
                Report.Reportx1.Value[0] = txtInvoiceNo.Text;
                Report.Reportx1 op = new Report.Reportx1("ExInvoiceTAX2.rpt");
            }
            catch { }
        }
    }
}
