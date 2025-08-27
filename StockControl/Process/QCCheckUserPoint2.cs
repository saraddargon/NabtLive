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
namespace StockControl
{
    public partial class QCCheckUserPoint2 : Telerik.WinControls.UI.RadRibbonForm
    {
        public QCCheckUserPoint2()
        {
            InitializeComponent();
        }
     
        public QCCheckUserPoint2(string QCNox)
        {
            InitializeComponent();
            QCNo = QCNox;
        }
     
        TextBox LinkPage = new TextBox();
        string Link = "";
        string QCNo = "";
        int id = 0;
        //private int RowView = 50;
        //private int ColView = 10;
        int AA = 0;
        RadTextBox NGidList = new RadTextBox();
     
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
            // RMenu3.Click += RMenu3_Click;
            // RMenu4.Click += RMenu4_Click;
            // RMenu5.Click += RMenu5_Click;
            // RMenu6.Click += RMenu6_Click;
            //// radGridView1.ReadOnly = true;
            // radGridView1.AutoGenerateColumns = false;
            // GETDTRow();
            cboStep.Items.Clear();
            cboStep.Items.Add("ประกอบ");
            cboStep.Items.Add("Test Leak");
            cboStep.Items.Add("ท้ายไลน์");

            cboTime.Items.Clear();
            cboTime.Items.Add("11:00-12:00");
            cboTime.Items.Add("12:00-13:00");
            cboTime.Items.Add("17:00-17:30");
            cboTime.Items.Add("17:30-18:00");

            cboTime.Items.Add("23:00-24:00");
            cboTime.Items.Add("00:00-01:00");
            cboTime.Items.Add("05:00-05:30");
            cboTime.Items.Add("05:30-06:00");



            DateTime date = DateTime.Now;
            var firstDayOfMonth = new DateTime(date.Year, date.Month, 1);
            var lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);
            //dtDate1.Value = firstDayOfMonth;
            //dtDate2.Value = lastDayOfMonth;
            DataLoad();
        

        }

        private void RMenu6_Click(object sender, EventArgs e)
        {
           
            DeleteUnit();
            DataLoad();
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

        private void DataLoad()
        {
            try
            {
                
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    radGridView1.DataSource = null;
                    var ListR = db.tb_QCCheckUserTimes.Where(p => p.QCNo.Equals(QCNo)).ToList();
                    if (ListR.Count > 0)
                        radGridView1.DataSource = ListR;


                }
            }
            catch { }

        }
        private bool CheckDuplicate(string code)
        {
            bool ck = false;


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
            // DataLoad();
            DataLoad();
        }
        private void NewClick()
        {
            //radGridView1.ReadOnly = false;
            //radGridView1.AllowAddNewRow = false;
            ////btnEdit.Enabled = false;
            //btnView.Enabled = true;
            //radGridView1.Rows.AddNew();
        }
        private void EditClick()
        {
            //radGridView1.ReadOnly = false;
            ////btnEdit.Enabled = false;
            //btnView.Enabled = true;
            //radGridView1.AllowAddNewRow = false;
        }
        private void ViewClick()
        {
           //// radGridView1.ReadOnly = true;
           // btnView.Enabled = false;
           // //btnEdit.Enabled = true;
           
            DataLoad();
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            NewClick();
        }

        private void btnView_Click(object sender, EventArgs e)
        {
            //Select NCR List No."//
            try
            {
                Saveclick();
            }
            catch { }
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {

            EditClick();
        }
        private void Saveclick()
        {
            if (MessageBox.Show("ต้องการบันทึก ?", "บันทึก", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                //AddUnit();
                //DataLoad();
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    tb_QCNGPoint ng = db.tb_QCNGPoints.Where(w => w.id.Equals(id)).FirstOrDefault();
                    if (ng != null)
                    {
                       
                        db.SubmitChanges();

                    }
                    else
                    {
                        
                      
                       

                    }
                    MessageBox.Show("บันทึกสำเร็จ");
                    DataLoad();
                }
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

        }

        private void Unit_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {


        }

        private void radGridView1_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {

            //if (e.KeyData == (Keys.Control | Keys.S))
            //{
            //    if (MessageBox.Show("ต้องการบันทึก ?", "บันทึก", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            //    {
            //        AddUnit();
            //        DataLoad();
            //    }
            //}
            //else if (e.KeyData == (Keys.Control | Keys.N))
            //{
            //    if (MessageBox.Show("ต้องการสร้างใหม่ ?", "สร้างใหม่", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            //    {
            //        NewClick();
            //    }
            //}

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
           
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
          
        }

        private void ImportData()
        {
          
        }

        private void btnFilter1_Click(object sender, EventArgs e)
        {
           
        }

        private void btnUnfilter1_Click(object sender, EventArgs e)
        {
            
        }

        private void radMenuItem1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            DataLoad();
        }

        private void radGridView1_CellDoubleClick(object sender, Telerik.WinControls.UI.GridViewCellEventArgs e)
        {
            try
            {
                //string RC = radGridView1.Rows[e.RowIndex].Cells["RCNo"].Value.ToString();
                //Receive rc = new Receive(RC);
                //rc.ShowDialog();
                // DataLoad();
                row = e.RowIndex;
            }
            catch { }
        }

        private void radButtonElement1_Click(object sender, EventArgs e)
        {
            try
            {
               

            }
            catch { }
        }

        private void radButtonElement2_Click(object sender, EventArgs e)
        {
            ReqMoveStock rq = new ReqMoveStock("", QCNo);
            rq.Show();
        }

        private void radButtonElement3_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Create NCR No.");
        }

        private void radButtonElement4_Click(object sender, EventArgs e)
        {

        }

        private void radGridView1_CellClick_1(object sender, GridViewCellEventArgs e)
        {
            row = e.RowIndex;
        }

        private void radButton1_Click(object sender, EventArgs e)
        {
            try
            {
                if (!txtChangeBy.Text.Equals(""))
                {
                    using (DataClasses1DataContext db = new DataClasses1DataContext())
                    {
                        tb_QCCheckUserTime ck = db.tb_QCCheckUserTimes.Where(p => p.UDesc.Equals(cboStep.Text) && p.BoxNo.Equals(cboTime.Text) && p.QCNo.Equals(QCNo)).FirstOrDefault();
                        if (ck == null)
                        {
                            tb_QCCheckUserTime qc = new tb_QCCheckUserTime();
                            qc.QCNo = QCNo;
                            qc.UserName = txtChangeBy.Text;
                            qc.DayN = "";
                            qc.UDesc = cboStep.Text;
                            qc.BoxNo = cboTime.Text;
                            qc.UType = "time";
                            qc.UserID = txtChangeBy.Text;
                            qc.ScanDate = DateTime.Now;
                            db.tb_QCCheckUserTimes.InsertOnSubmit(qc);
                            db.SubmitChanges();
                        }

                    }
                }
            }
            catch { }
            DataLoad();
            txtChangeBy.Text = "";
        }

        private void radButtonElement5_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("ต้องการลบหรือไม่ ?", "delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                int id = 0;
                int.TryParse(radGridView1.CurrentRow.Cells["id"].Value.ToString(), out id);
                if (id > 0)
                {
                    using (DataClasses1DataContext db = new DataClasses1DataContext())
                    {
                        tb_QCCheckUserTime dl = db.tb_QCCheckUserTimes.Where(p => p.id.Equals(id)).FirstOrDefault();
                        if(dl!=null)
                        {
                            db.tb_QCCheckUserTimes.DeleteOnSubmit(dl);
                            db.SubmitChanges();
                            MessageBox.Show("Delete Completed.");
                            DataLoad();
                        }
                    }
                }
            }
        }
    }
}
