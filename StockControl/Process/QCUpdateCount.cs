using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Linq;
using Microsoft.VisualBasic.FileIO;
using System.Runtime.InteropServices;
namespace StockControl
{
    public partial class QCUpdateCount : Telerik.WinControls.UI.RadRibbonForm
    {
   
        Telerik.WinControls.UI.RadTextBox CodeNo_tt = new Telerik.WinControls.UI.RadTextBox();
        int screen = 0;
        public QCUpdateCount(string  CodeNox)
        {
            InitializeComponent();            
            screen = 1;
            txtWoNo.Text = CodeNox;
        }
        public QCUpdateCount()
        {
            InitializeComponent();
        }
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
           
            if (keyData == Keys.Up)
            {
                UpDownData(-1);
            }
            else if (keyData == Keys.Down)
            {
                UpDownData(1);
            }


            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void UpDownData(int a)
        {
            try
            {
                Gobal += a;
                if (Gobal <= 0)
                    Gobal = 1;
                if (Gobal > 10)
                    Gobal = 10;
                int QQ = 0;
                lblseq.Text = Gobal.ToString();
                ///////Show data////
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    tb_QCCountPD qp = db.tb_QCCountPDs.Where(c => c.WONo.Equals(txtWoNo.Text) && c.Seq.Equals(Gobal)).FirstOrDefault();
                    if(qp!=null)
                    {
                        QQ = Convert.ToInt32(qp.A1);
                    }
                }
                txtScan.Text = ValueText(Gobal);
                txtQty.Text = QQ.ToString();
                txtQty.Focus();
                //if (Gobal.Equals(1))
                //{
                //    txtScan.Text = ValueText(1);
                //    txtQty.Text = QQ.ToString();
                //    txtQty.Focus();

                //}
                //else if (Gobal.Equals(2))
                //{
                //    txtScan.Text = ValueText(2);
                //    txtQty.Text = QQ.ToString();
                //    txtQty.Focus();
                //}
                //else if (Gobal.Equals(3))
                //{
                //    txtScan.Text = ValueText(3);
                //    txtQty.Text = QQ.ToString();
                //    txtQty.Focus();
                //}
                //else if (Gobal.Equals(4))
                //{
                //    txtScan.Text = ValueText(3);
                //    txtQty.Text = QQ.ToString();
                //    txtQty.Focus();
                //}
                //else if (Gobal.Equals(5))
                //{
                //    txtScan.Text = "เครื่องอัดสปริง";
                //    txtQty.Text = QQ.ToString();
                //    txtQty.Focus();
                //}
                //else if (Gobal.Equals(6))
                //{
                //    txtScan.Text = "เครื่องรัดขอบ";
                //    txtQty.Text = QQ.ToString();
                //    txtQty.Focus();
                //}
                //else if (Gobal.Equals(7))
                //{
                //    txtScan.Text = "เครื่องพ่นสี";
                //    txtQty.Text = QQ.ToString();
                //    txtQty.Focus();
                //}
                //else if (Gobal.Equals(8))
                //{
                //    txtScan.Text = "เครื่อง Test";
                //    txtQty.Text = QQ.ToString();
                //    txtQty.Focus();

                //}
                //else if (Gobal.Equals(9))
                //{
                //    txtScan.Text = "เครื่อง Stamp Lot";
                //    txtQty.Text = QQ.ToString();
                //    txtQty.Focus();
                //}
                //else if (Gobal.Equals(10))
                //{
                //    txtScan.Text = "ท้ายไลน์";
                //    txtQty.Text = QQ.ToString();
                //    txtQty.Focus();
                //}



            }
            catch { }
        }
        int Gobal = 0;
        private string ValueText(int aa)
        {
            string TX = "";
            if (aa.Equals(1))
            {
                TX = "SUB :Seal and Guide";
                

            }
            else if (aa.Equals(2))
            {
                TX = "SUB :Pre Tighten";
               
            }
            else if (aa.Equals(3))
            {
                TX = "SUB:Case Assemble";
                
            }
            else if (aa.Equals(4))
            {
                TX = "เครื่องลอกสี";
                
            }
            else if (aa.Equals(5))
            {
                TX = "เครื่องอัดสปริง";
               
            }
            else if (aa.Equals(6))
            {
                TX = "เครื่องรัดขอบ";
                
            }
            else if (aa.Equals(7))
            {
                TX = "เครื่องพ่นสี";
               
            }
            else if (aa.Equals(8))
            {
                TX = "เครื่อง Test";
                

            }
            else if (aa.Equals(9))
            {
                TX = "เครื่อง Stamp Lot";              
            }
            else if (aa.Equals(10))
            {
                TX = "ท้ายไลน์";
               
            }
            return TX;
        }
        //private int RowView = 50;
        //private int ColView = 10;
        //DataTable dt = new DataTable();
        private void radMenuItem2_Click(object sender, EventArgs e)
        {

        }

        private void radRibbonBar1_Click(object sender, EventArgs e)
        {

        }
        private void clearDb()
        {
            try
            {
                //txtA1.Text = "";
                //txtA2.Text = "";
                //txtA3.Text = "";
                //txtA4.Text = "";
                //txtA5.Text = "";
                //txtA1N.Text = "";
                //txtA2N.Text = "";
                //txtA3N.Text = "";
                //txtA4N.Text = "";
                //txtA5N.Text = "";

                //txtB1.Text = "";
                //txtB2.Text = "";
                //txtB3.Text = "";
                //txtB4.Text = "";
                //txtB5.Text = "";
              

                //txtB1N.Text = "";
                //txtB2N.Text = "";
                //txtB3N.Text = "";
                //txtB4N.Text = "";
                //txtB5N.Text = "";

            }
            catch { }
        }
        private void Unit_Load(object sender, EventArgs e)
        {
            lblseq.Text = Gobal.ToString();
            clearDb();
            LoadData();
            UpDownData(1);


        }
        private void LoadData()
        {
            using (DataClasses1DataContext db = new DataClasses1DataContext())
            {
                var ListB = db.tb_QCCountPDs.Where(c => c.WONo.Equals(txtWoNo.Text)).ToList();
                radGridView4.DataSource = null;
                radGridView4.DataSource = ListB;               


            }
        }

        private void btn_PrintPR_Click(object sender, EventArgs e)
        {

        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            if(MessageBox.Show("ต้องการบันทึกหรือไม่ ?","การบันทึก",MessageBoxButtons.YesNo,MessageBoxIcon.Question)==DialogResult.Yes)
            {
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                   // db.sp_46_QCUpdateLot(txtWoNo.Text, txtLot.Text, rdoWorkShift.Text);
                    MessageBox.Show("บันทึกแล้ว");
                }
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            UpDownData(-1);
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            LoadData();
        }

        private void radButton1_Click(object sender, EventArgs e)
        {
            UpDownData(1);
        }

        private void txtQty_KeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {
               if(e.KeyChar==13)
                {
                    SaveData();
                }
            }
            catch { }
        }
        private void SaveData()
        {
            try
            {
                if(Gobal>0)
                {
                    using (DataClasses1DataContext db = new DataClasses1DataContext())
                    {
                        string DN = dbShowData.CheckDayN(DateTime.Now);
                        int qq = 0;
                        int.TryParse(txtQty.Text, out qq);

                        tb_QCCountPD ck = db.tb_QCCountPDs.Where(c => c.WONo.Equals(txtWoNo.Text) && c.Seq.Equals(Gobal) && c.DayN.Equals(DN)).FirstOrDefault();
                        if(ck!=null)
                        {
                            ck.A1 = qq;
                            db.SubmitChanges();
                        }else
                        {
                            tb_QCCountPD nc = new tb_QCCountPD();
                            nc.WONo = txtWoNo.Text.ToUpper();
                            nc.DayN = DN;
                            nc.A1 = qq;
                            nc.ProcessName = ValueText(Gobal);
                            nc.Seq = Gobal;
                            db.tb_QCCountPDs.InsertOnSubmit(nc);
                            db.SubmitChanges();
                        }

                    }
                }
            }
            catch { }
            LoadData();
            txtQty.Focus();

        }
        int rows = 0;
        private void radGridView4_CellClick(object sender, Telerik.WinControls.UI.GridViewCellEventArgs e)
        {
            rows = e.RowIndex;
        }

        private void deleteLineToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (rows >= 0)
                {
                    int seq = Convert.ToInt32(radGridView4.Rows[rows].Cells["Seq"].Value.ToString());
                    if (seq > 0)
                    {
                        if (MessageBox.Show("ต้องการลบ หรือไม่ ?", "ลบรายการ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            using (DataClasses1DataContext db = new DataClasses1DataContext())
                            {
                                tb_QCCountPD dl = db.tb_QCCountPDs.Where(c => c.WONo.Equals(txtWoNo.Text) && c.Seq.Equals(seq)).FirstOrDefault();
                                if (dl != null)
                                {
                                    db.tb_QCCountPDs.DeleteOnSubmit(dl);
                                    db.SubmitChanges();
                                    MessageBox.Show("ลบรายการแล้ว");
                                    LoadData();
                                }
                            }
                        }
                    }
                }
            }
            catch { }
        }

        private void radPanel1_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
