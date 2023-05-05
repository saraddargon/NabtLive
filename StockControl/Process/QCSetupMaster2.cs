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

namespace StockControl
{
    public partial class QCSetupMaster2 : Telerik.WinControls.UI.RadRibbonForm
    {
        public QCSetupMaster2()
        {
            InitializeComponent();
        }
        public QCSetupMaster2(string codex)
        {
            InitializeComponent();
            Code = codex;
            txtScanID.Text = Code;
            
        }
        string Code = "";
        string PType = "";
        string pathfile = "";
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
        private void GETDTRow()
        {
            //dt.Columns.Add(new DataColumn("UnitCode", typeof(string)));
            //dt.Columns.Add(new DataColumn("UnitDetail", typeof(string)));
            //dt.Columns.Add(new DataColumn("UnitActive", typeof(bool)));
        }
        private void Unit_Load(object sender, EventArgs e)
        {
            setDefault();
            DataLoad();
        }
        private void setDefault()
        {
            dtDate.Value = DateTime.Now;
            using (DataClasses1DataContext db = new DataClasses1DataContext())
            {
                pathfile = "";
                tb_Path ph = db.tb_Paths.Where(p => p.PathCode.Equals("TemplateQC")).FirstOrDefault();
                if (ph != null)
                {
                    pathfile = ph.PathFile;
                }
                var qf = db.tb_QCFormMasters.ToList();
                cboISO.DataSource = null;
                cboISO.DataSource = qf;
                cboISO.ValueMember = "FormISO";
                cboISO.DisplayMember = "FormISO";
            }
        }
        private void DataLoad()
        {
            try
            {
                radGridView2.DataSource = null;
                radGridView2.AutoGenerateColumns = false;
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    radGridView2.DataSource = db.tb_QCFormStartDates.Where(p => p.PartNo.Equals(Code)).OrderBy(s=>s.FormISO).OrderBy(s2=>s2.Version).ToList();
                }
            }
            catch { }
        }
        private void SetFocus()
        {
           
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
            DataLoad();

        }
        private void NewClick()
        {
          
        }
        private void EditClick()
        {
          
        }
        private void ViewClick()
        {
         
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
           
        }
        private void UploadImage(string Path,string Listpath)
        {
          
        }
        private void DeleteClick()
        {

        }

        private void btnSave_Click(object sender, EventArgs e)
        {
         
        }


        private void radGridView1_CellEndEdit(object sender, Telerik.WinControls.UI.GridViewCellEventArgs e)
        {
           
        }

        private void Unit_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {


        }

        private void radGridView1_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {

            if (e.KeyData == (Keys.Control | Keys.S))
            {
                //if (MessageBox.Show("ต้องการบันทึก ?", "บันทึก", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                //{
                //    //AddUnit();
                //    //DataLoad();
                //}
            }
            else if (e.KeyData == (Keys.Control | Keys.N))
            {
                //if (MessageBox.Show("ต้องการสร้างใหม่ ?", "สร้างใหม่", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                //{
                //    //NewClick();
                //}
            }

        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            
        
            
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


        private void ImportData()
        {
           
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

        private void btnImage_Click(object sender, EventArgs e)
        {
            
        }

        private void radButton1_Click(object sender, EventArgs e)
        {
           
        }

        private void radButton2_Click(object sender, EventArgs e)
        {
            
        }

        private void radButton3_Click(object sender, EventArgs e)
        {
            
        }

        private void txtScanID_KeyPress(object sender, KeyPressEventArgs e)
        {

        }
        string PDTAG = "";
       

        private void radGridView2_CellDoubleClick(object sender, Telerik.WinControls.UI.GridViewCellEventArgs e)
        {

        }

        private void radButton2_Click_1(object sender, EventArgs e)
        {
            //Seleect File//
            OpenFileDialog op = new OpenFileDialog();
            op.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.xlsx)|*.xlsx";
            if (op.ShowDialog() == DialogResult.OK)
            {
                txtFilePath.Text = op.FileName ;
                ImportData(op.FileName);
            }
        }
        private void ImportData(string FileName)
        {
            try
            {

            }
            catch { }
        }

        private void radButtonElement2_Click(object sender, EventArgs e)
        {
            try
            {
                if(radGridView2.CurrentRow.Index>=0)
                {
                    string fname = radGridView2.CurrentRow.Cells["FileName"].Value.ToString();
                    if(pathfile!="")
                    {
                        string tempPath = System.IO.Path.GetTempPath();
                        string tempfile = tempPath + fname;
                        string Source = pathfile + fname;
                        if (File.Exists(tempfile))
                        {
                            try
                            {
                                File.Delete(tempfile);
                            }
                            catch { }
                        }
                        else
                        {
                            
                            File.Copy(Source, tempfile);
                            File.Open(tempfile, FileMode.Open);
                        }

                    }

                }
            }
            catch(Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void radButtonElement1_Click(object sender, EventArgs e)
        {
            try
            {
                if(MessageBox.Show("ต้องการลบรายการหรือไม่ ?","ลบรายการ",MessageBoxButtons.YesNo,MessageBoxIcon.Question)==DialogResult.Yes)
                {
                    if (radGridView2.CurrentRow.Index >= 0)
                    {
                        string fname = radGridView2.CurrentRow.Cells["FileName"].Value.ToString();
                        int idx = Convert.ToInt32(radGridView2.CurrentRow.Cells["id"].Value.ToString());
                        if (pathfile != "" && idx>0)
                        {
                            using (DataClasses1DataContext db = new DataClasses1DataContext())
                            {
                                tb_QCFormStartDate qcf = db.tb_QCFormStartDates.Where(p => p.id.Equals(idx)).FirstOrDefault();
                                if(qcf!=null)
                                {
                                    db.tb_QCFormStartDates.DeleteOnSubmit(qcf);
                                    db.SubmitChanges();
                                }
                            }
                        
                            string Source = pathfile + fname;
                            if (File.Exists(Source))
                            {
                                try
                                {
                                    File.Delete(Source);
                                }
                                catch { }
                            }
                           

                        }
                    }
                }
            }
            catch { }
        }

        private void radButton1_Click_1(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("ต้องการบันทึกหรือไม่ ?", "บันทึก", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    if (!txtScanID.Text.Equals("") && !txtFilePath.Text.Equals("") && !txtRev.Text.Equals("") && !cboISO.Text.Equals(""))
                    {
                        string fileN = Path.GetFileName(txtFilePath.Text);
                        using (DataClasses1DataContext db = new DataClasses1DataContext())
                        {
                            tb_QCFormStartDate cke = db.tb_QCFormStartDates.Where(p => p.FileName.Equals(fileN)).FirstOrDefault();
                            if(cke!=null)
                            {
                                MessageBox.Show("ชื่อไฟล์ซ้ำ!!!");
                            }
                            else
                            {
                                //Add
                                tb_QCFormStartDate qcadd = new tb_QCFormStartDate();
                                qcadd.FormISO = cboISO.Text;
                                qcadd.FileName = fileN;
                                qcadd.StartDate = dtDate.Value;
                                qcadd.PartNo = txtScanID.Text;
                                qcadd.Version = txtRev.Text;
                                db.tb_QCFormStartDates.InsertOnSubmit(qcadd);
                                db.SubmitChanges();
                                DataLoad();
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("ต้องใส่ค่าให้ครบทุกช่อง");
                    }
                }
            }
            catch { }
        }
    }
}
