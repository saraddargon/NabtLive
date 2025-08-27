using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Linq;
using Telerik.WinControls.UI;
using Microsoft.VisualBasic.FileIO;
using System.Runtime.InteropServices;
namespace StockControl
{
    public partial class QCSetMasterSelect : Telerik.WinControls.UI.RadRibbonForm
    {
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
           // MessageBox.Show(keyData.ToString());
            if (keyData == (Keys.Control | Keys.S))
            {
                // Alt+F pressed
                //  ClearData();

                return false;
                //txtSeriesNo.Focus();
            }
            else if ((keyData == Keys.NumPad1 ) || (keyData== Keys.D1))
            {
                radioButton1.Checked = true;
            }
            else if ((keyData == Keys.NumPad2) || (keyData == Keys.D2))
            {
                radioButton2.Checked = true;
            }
            else if ((keyData == Keys.NumPad3) || (keyData == Keys.D3))
            {
                radioButton3.Checked = true;
            }
            else if ((keyData == Keys.NumPad4) || (keyData == Keys.D4))
            {
                radioButton4.Checked = true;
            }
            else if ((keyData == Keys.NumPad5) || (keyData == Keys.D5))
            {
                radioButton5.Checked = true;
            }
            else if ((keyData == Keys.NumPad6) || (keyData == Keys.D6))
            {
                radioButton6.Checked = true;
            }
            else if (keyData == (Keys.F9))
            {
                SelectLoad();
            }
            else if (keyData == (Keys.Escape))
            {
                this.Close();
            }
           

            return base.ProcessCmdKey(ref msg, keyData);
        }

        public QCSetMasterSelect()
        {
            InitializeComponent();
        }
        public QCSetMasterSelect(string OrderNo,string Linex,string PartNo,RadTextBox tx,string Ty)
        {
            InitializeComponent();
            WONo = OrderNo;
            LineName = Linex;
            Code = PartNo;
            ISO = tx;
            PType = Ty;
        }
        string Code = "";
        string PType = "";
        string LineName = "";
        string WONo = "";
        RadTextBox ISO = new RadTextBox();
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
        private void Unit_Load(object sender, EventArgs e)
        {

            try
            {
                radioButton1.Visible = false;
                radioButton2.Visible = false;
                radioButton3.Visible = false;
                radioButton4.Visible = false;
                radioButton5.Visible = false;
                radioButton6.Visible = false;
                int ck = 0;
                txtPartNo.Text = Code;

                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    //  var listQ= db.sp_46_QCSelectWO_02(WONo, LineName, Code, PType).ToList();
                    var listQ = db.sp_46_QCSelectWO_02xV1(Code, PType).ToList();
                    int CountA = 0;
                    foreach(var rd in listQ)
                    {
                        CountA += 1;
                        if (CountA == 1)
                        {
                            radioButton1.Text = rd.FormISO + " " + rd.FormName;
                            radTextBox1.Text = rd.FormISO;
                            radioButton1.Visible = true;
                            SetVisible(rd.FormISO,ref radioButton1);
                        }
                        else if (CountA == 2)
                        {
                            radioButton2.Text = rd.FormISO + " " + rd.FormName;
                            radTextBox2.Text = rd.FormISO;
                            radioButton2.Visible = true;
                            SetVisible(rd.FormISO, ref radioButton2);
                        }
                        else if (CountA == 3)
                        {
                            radioButton3.Text = rd.FormISO + " " + rd.FormName;
                            radTextBox3.Text = rd.FormISO;
                            radioButton3.Visible = true;
                            SetVisible(rd.FormISO, ref radioButton3);
                        }
                        else if (CountA == 4)
                        {
                            radioButton4.Text = rd.FormISO + " " + rd.FormName;
                            radTextBox4.Text = rd.FormISO;
                            radioButton4.Visible = true;
                            SetVisible(rd.FormISO, ref radioButton4);
                        }
                        else if (CountA == 5)
                        {
                            radioButton5.Text = rd.FormISO + " " + rd.FormName;
                            radTextBox5.Text = rd.FormISO;
                            radioButton5.Visible = true;
                            SetVisible(rd.FormISO, ref radioButton5);
                        }
                        else if (CountA == 6)
                        {
                            radioButton6.Text = rd.FormISO + " " + rd.FormName;
                            radTextBox6.Text = rd.FormISO;
                            radioButton6.Visible = true;
                            SetVisible(rd.FormISO, ref radioButton6);
                        }
                    
                    }
                }



            }
            catch(Exception ex) { MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }
        private void SetVisible(string FormISO, ref RadioButton btn1)
        {
            ///////////

            if (FormISO.ToString().Equals("FM-QA-091"))
            {
                btn1.Visible = false;
                if (txtPartNo.Text.Equals("37116013410") || txtPartNo.Text.Equals("37116013780") || txtPartNo.Text.Equals("37116013230"))
                {
                    btn1.Visible = true;
                    //ck += 1;
                }
            }
            if (FormISO.ToString().Equals("FM-QA-092"))
            {
                btn1.Visible = false;
                if (txtPartNo.Text.Equals("37100013410") || txtPartNo.Text.Equals("37100013420") || txtPartNo.Text.Equals("37100013780")
                    || txtPartNo.Text.Equals("37100013431") || txtPartNo.Text.Equals("37100013680") || txtPartNo.Text.Equals("31AZ-032331M")
                    || txtPartNo.Text.Equals("37100013780A5"))
                {
                    btn1.Visible = true;
                    // ck += 1;
                }
                //New
                if (txtPartNo.Text.Equals("37100013410S") || txtPartNo.Text.Equals("37100013420S") || txtPartNo.Text.Equals("37100013780S")
                  || txtPartNo.Text.Equals("37100013431S") || txtPartNo.Text.Equals("37100013680S"))
                {
                    btn1.Visible = true;
                    // ck += 1;
                }
            }
            if (FormISO.ToString().Equals("FM-QA-143"))
            {
                btn1.Visible = false;
                if (txtPartNo.Text.Equals("37116013690"))
                {
                    btn1.Visible = true;
                    // ck += 1;
                }
            }
            if (FormISO.ToString().Equals("FM-QA-144"))
            {
                btn1.Visible = false;
                if (txtPartNo.Text.Equals("37100013690") || txtPartNo.Text.Equals("37100013690T") || txtPartNo.Text.Equals("37100013770")
                    || txtPartNo.Text.Equals("37100013690S") || txtPartNo.Text.Equals("37100013770S"))
                {
                    btn1.Visible = true;
                    // ck += 1;
                }
            }
            if (FormISO.ToString().Equals("FM-QA-161"))
            {
                btn1.Visible = false;
                if (txtPartNo.Text.Equals("37200014141") || txtPartNo.Text.Equals("37200014151") || txtPartNo.Text.Equals("37200014161"))
                {
                    btn1.Visible = true;
                    // ck += 1;
                }
            }

            ////new
            if (FormISO.ToString().Equals("FM-QA-055") || FormISO.Equals("FM-QA-055_02_1"))
            {
                string FormS = "FM-QA-055";
                
                btn1.Visible = true;
                //if (txtPartNo.Text.Equals("37116013690")
                //    || txtPartNo.Text.Equals("37116013230")
                //    || txtPartNo.Text.Equals("37116013410")
                //    || txtPartNo.Text.Equals("37116013780"))
                //{
                //    btn1.Visible = false;
                //    // ck += 1;
                //}
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    if(Convert.ToInt32(db.get_QCSkip(FormS,txtPartNo.Text,0))>0)
                    {
                        btn1.Visible = false;
                    }
                }

            }
            if (FormISO.ToString().Equals("FM-PD-011") || (FormISO.ToString().Equals("FM-PD-010")))
            {
                btn1.Visible = true;
                if (LineName.Equals("TD11-DR SUB 1")
                    || LineName.Equals("TD12-DR SUB-2")
                     || LineName.Equals("TD13-DR SUB-3")
                      || LineName.Equals("TD14-DR SUB-4")
                      || LineName.Equals("TD15-DR SUB-5")
                      || LineName.Equals("TD16-DR SUB-6")
                      || LineName.Equals("TD17-DR SUB-2")
                    )
                {
                    btn1.Visible = false;
                }

            }
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

        private void DataLoad()
        {
           
            
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
            //DataLoad();
           

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
            SelectLoad();
        }
        private void SelectLoad()
        {
            if (radioButton1.Checked)
            {
                ISO.Text = radTextBox1.Text;
            }
            else if (radioButton2.Checked)
            {
                ISO.Text = radTextBox2.Text;
            }
            else if (radioButton3.Checked)
            {
                ISO.Text = radTextBox3.Text;
            }
            else if (radioButton4.Checked)
            {
                ISO.Text = radTextBox4.Text;
            }
            else if (radioButton5.Checked)
            {
                ISO.Text = radTextBox5.Text;
            }
            else if (radioButton6.Checked)
            {
                ISO.Text = radTextBox6.Text;
            }

            this.Close();
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
            
                DeleteUnit();
                DataLoad();
            
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
            if(e.KeyChar==13)
            {
                getWO();
            }
        }
        string PDTAG = "";
        private void getWO()
        {
                      
        }

        private void radGridView2_CellDoubleClick(object sender, Telerik.WinControls.UI.GridViewCellEventArgs e)
        {
           
        }
    }
}
