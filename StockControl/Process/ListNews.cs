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
    public partial class ListNews : Telerik.WinControls.UI.RadRibbonForm
    {
        public ListNews(string ACC)
        {
            this.Name = "ListNews";
            //if (!dbClss.PermissionScreen(this.Name))
            //{
            //    MessageBox.Show("Access denied", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    this.Close();
            //}
           
            AC = ACC;
           
            InitializeComponent();
            this.Text = "News & Forcast";
            lblType.Text = AC;
        }
        string AC = "";
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
            radGridView1.ReadOnly = true;
            radGridView1.AutoGenerateColumns = false; 
            DataLoad();
        }

        private void RMenu6_Click(object sender, EventArgs e)
        {
           
            //DeleteUnit();
            //DataLoad();
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
                this.Cursor = Cursors.WaitCursor;
                int ck = 0;
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {


                    radGridView1.DataSource = db.sp_60_Select_NewsForcast(AC).ToList();
                    foreach (var x in radGridView1.Rows)
                    {                        
                        ck += 1;
                        x.Cells["No"].Value = ck;
                    }

                }
            }
            catch { }
            this.Cursor = Cursors.Default;



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
        
            DataLoad();
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            NewClick();
        }

        private void btnView_Click(object sender, EventArgs e)
        {
            ViewClick();
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {

            EditClick();
        }
        private void Saveclick()
        {
           
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
           dbClss.ExportGridXlSX(radGridView1);
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
     
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

        private void btnSave_Click_1(object sender, EventArgs e)
        {
            //Example01.pdf
         //   System.Diagnostics.Process.Start(Environment.CurrentDirectory+@"\Example\Example01.pdf");

        }

        private void radButton1_Click(object sender, EventArgs e)
        {
            DataLoad();
        }

        private void chkSelect_ToggleStateChanged(object sender, Telerik.WinControls.UI.StateChangedEventArgs args)
        {
           
        }

        private void btnSave_Click_2(object sender, EventArgs e)
        {
            try
            {
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    if (MessageBox.Show("ต้องการ ลบ", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        int id = 0;
                        id = Convert.ToInt32(radGridView1.CurrentRow.Cells["id"].Value.ToString());
                        if(id>0)
                        {
                            db.sp_60_Delete_NewsForcast(id);
                            DataLoad();
                        }
                    }
                }
            }
            catch { }

        }
       

        private void txtItemNo_KeyPress(object sender, KeyPressEventArgs e)
        {
            if(e.KeyChar==13)
            {
                DataLoad();
            }
        }

        private void radButtonElement1_Click(object sender, EventArgs e)
        {
          
        }

        private void radButton2_Click(object sender, EventArgs e)
        {
            using (DataClasses1DataContext db = new DataClasses1DataContext())
            {
                if (!txtFile.Text.Equals(""))
                {
                    if (MessageBox.Show("ต้องการบันทึก", "บันทึก", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        string path = db.web_getPath();
                        string Ext = System.IO.Path.GetExtension(txtFile.Text);
                        string FileName = AC + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ""+ Ext;
                        if (!path.Equals(""))
                        {
                            System.IO.File.Copy(txtFile.Text,path + FileName, true);
                            db.sp_60_AddNewForcast(AC, txtTopic.Text, txtDetail.Text, txtRemark.Text, FileName, dbClss.UserID,txtVendorNo.Text);
                            MessageBox.Show("Completed.");
                            DataLoad();
                        }
                        else
                        {
                            MessageBox.Show("Path Invalid!");
                        }
                    }
                }
            }
        }

        private void radButton1_Click_1(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "files (*.pdf;*.xlsx)|*.pdf;*.xlsx";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.FileName = "";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {

                txtFile.Text = openFileDialog1.FileName;
            }
        }

        private void radButton3_Click(object sender, EventArgs e)
        {
            txtDetail.Text = "";
            txtFile.Text = "";
            txtRemark.Text = "";
            txtTopic.Text = "";
        }
    }
}
