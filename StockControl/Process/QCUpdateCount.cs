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
        public QCUpdateCount(string  CodeNox,string LineName,string PartNox)
        {
            InitializeComponent();            
            screen = 1;
            txtWoNo.Text = CodeNox;
            txtLineName.Text = LineName;
            PartNo = PartNox;
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
        string TempReport = "";
        string PartNo = "";
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

                    
                }//TW10-CB
                if (txtLineName.Text.Equals("TW10-CB") || txtLineName.Text.Equals("TW20-CB"))
                {
                    txtScan.Text = ValueText2(Gobal);
                }
                else if(txtLineName.Text.Equals("TW02-SC_PB"))
                {
                    txtScan.Text = ValueText3(Gobal);
                }
                else if(txtLineName.Text.Equals("TS10-SC_CB"))                    
                {                    
                    txtScan.Text = ValueText4(Gobal);
                }
                else if(txtLineName.Text.Equals("TC10-CL-MT AAT")
                    || txtLineName.Text.Equals("TC20-MAIN_M")
                    || txtLineName.Text.Equals("TC30-MAIN_T"))
                {
                    if (TempReport.Equals("NISSAN"))
                    {
                        txtScan.Text = ValueText5B(Gobal);
                    }
                    else if (TempReport.Equals("DATT"))
                    {
                        txtScan.Text = ValueText5A(Gobal);
                    }
                    else
                    {
                        txtScan.Text = ValueText5(Gobal);
                    }
                }
                else if(txtLineName.Text.Equals("TR10-RV-6"))
                {
                    txtScan.Text = ValueTextRV6(Gobal);
                }
                else if(txtLineName.Text.Equals("TP01-CL-OPE_SUB")
                    || txtLineName.Text.Equals("TP10-CL-OPE_MAIN")
                    )
                {
                    txtScan.Text = ValueText6A(Gobal);
                }
                else
                {
                    txtScan.Text = ValueText(Gobal);
                }
                txtQty.Text = QQ.ToString();
                txtQty.Focus();
               


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
                TX = "เครื่อง Test";                

            }
            else if (aa.Equals(8))
            {
                TX = "เครื่อง Stamp Lot";              
            }
            else if (aa.Equals(9))
            {
                TX = "เครื่องพ่นสี";

            }
            else if (aa.Equals(10))
            {
                TX = "ท้ายไลน์";
               
            }
            return TX;
        }
        private string ValueText2(int aa)
        {
            string TX = "";
            if (aa.Equals(1))
            {
                TX = "เครื่อง Sub Line";


            }
            else if (aa.Equals(2))
            {
                TX = "เครื่องตอก Lot No";

            }
            else if (aa.Equals(3))
            {
                TX = "เครื่องประกอบ";

            }
            else if (aa.Equals(4))
            {
                TX = "เครื่อง Tesk Leak (1)";

            }
            else if (aa.Equals(5))
            {
                TX = "เครื่อง Tesk Leak (2)";

            }
            else if (aa.Equals(6))
            {
                TX = "Check 100% Inspection";

            }          
            return TX;
        }
        private string ValueText3(int aa)
        {
            string TX = "";
            if (aa.Equals(1))
            {
                TX = "เครื่องลอกสี";
            }
            else if (aa.Equals(2))
            {
                TX = "Piston Rod";

            }
            else if (aa.Equals(3))
            {
                TX = "เครื่องอัดสปริง";

            }
            else if (aa.Equals(4))
            {
                TX = "เครื่องรัดขอบ";

            }
            else if (aa.Equals(5))
            {
                TX = "ประกอบ Nut";

            }
            ////P2
            else if (aa.Equals(6))
            {
                TX = "เครื่องพ่นสี";

            }
            else if (aa.Equals(7))
            {
                TX = "เครื่อง Test";

            }
            else if (aa.Equals(8))
            {
                TX = "เครื่อง Stamp Lot";

            }
            else if (aa.Equals(9))
            {
                TX = "ประกอบ Elbow Joint";

            }

            return TX;
        }
        private string ValueText4(int aa)
        {
            string TX = "";
            if (aa.Equals(1))
            {
                TX = "เครื่องStamp Clamp Ring";
            }
            else if (aa.Equals(2))
            {
                TX = "เครื่องประกอบ";

            }
            else if (aa.Equals(3))
            {
                TX = "เครื่อง Test 1";

            }
            else if (aa.Equals(4))
            {
                TX = "เครื่อง Test 2";

            }           

            return TX;
        }
        private string ValueText5(int aa)
        {
            string TX = "";
            if (aa.Equals(1))
            {
                TX = "Body Leak Tester (A2)";
            }
            //else if (aa.Equals(2))
            //{
            //    TX = "Piston Comp Assembly(A3)";

            //}
            else if (aa.Equals(2))
            {
                TX = "Performance Test (A4)";

            }
            else if (aa.Equals(3))
            {
                TX = "NIPPLE&RESERVOIR ASSEMBLY (A5)";

            }
            else if (aa.Equals(4))
            {
                TX = "NIPPLE&RESERVOIR Leak Tester (A6)";

            }
            else if (aa.Equals(5))
            {
                TX = "เครื่องขัน Stud Bolt(A7)";

            }
            else if (aa.Equals(6))
            {
                TX = "Final Inspection (A8)";

            }

            return TX;
        }
        private string ValueText5A(int aa)
        {
            string TX = "";
            if (aa.Equals(1))
            {
                TX = "Body Leak Tester (A2)";
            }
            //else if (aa.Equals(2))
            //{
            //    TX = "Piston Comp Assembly(A3)";

            //}
            else if (aa.Equals(2))
            {
                TX = "Performance Test (A4)";

            }
            else if (aa.Equals(3))
            {
                TX = "NIPPLE&RESERVOIR ASSEMBLY (A5)";

            }
            else if (aa.Equals(4))
            {
                TX = "NIPPLE&RESERVOIR Leak Tester (A6)";

            }
            else if (aa.Equals(5))
            {
                TX = "Final Inspection (A8)";

            }

            return TX;
        }
        private string ValueText5B(int aa)
        {
            string TX = "";
             if (aa.Equals(1))
            {
                TX = "Dcp Tighening (A1)";
            }
            else if (aa.Equals(2))
            {
                TX = "Body Leak Tester (A2)";
            }
            //else if (aa.Equals(3))
            //{
            //    TX = "Piston Comp Assembly(A3)";

            //}
            else if (aa.Equals(3))
            {
                TX = "Performance Test (A4)";

            }
            else if (aa.Equals(4))
            {
                TX = "NIPPLE&RESERVOIR ASSEMBLY (A5)";

            }
            else if (aa.Equals(5))
            {
                TX = "NIPPLE&RESERVOIR Leak Tester (A6)";
            }          
            else if (aa.Equals(6))
            {
                TX = "Final Inspection (A8)";

            }

            return TX;
        }
        private string ValueText6A(int aa)
        {
            string TX = "";
            if (aa.Equals(1))
            {
                TX = "Body Leak Tester (A2)";
            }
            else if (aa.Equals(2))
            {
                TX = "Assy Test (A5)";
            }
            else if (aa.Equals(3))
            {
                TX = "Lot Stamp (A8)";

            }
            else if (aa.Equals(4))
            {
                TX = "Final Inspection";

            }
          

            return TX;
        }
        private string ValueTextRV6(int aa)
        {
            string TX = "";
            if (aa.Equals(1))
            {
                TX = "Control Pressure &Leak Test  (A4)";
            }
            else if (aa.Equals(2))
            {
                TX = "Process  External Leak Test (A5)";
            }
            else if (aa.Equals(3))
            {
                TX = "Final Assembly Inspection (A6)";

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
            using (DataClasses1DataContext db = new DataClasses1DataContext())
            {
                TempReport = db.get_QC_SetDataMaster(db.get_QC_SetDataMaster_Line(txtLineName.Text), PartNo, 110);
            }

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
                            if (!txtScan.Text.Equals(""))
                            {
                                tb_QCCountPD nc = new tb_QCCountPD();
                                nc.WONo = txtWoNo.Text.ToUpper();
                                nc.DayN = DN;
                                nc.A1 = qq;

                                if (txtLineName.Text.Equals("TW10-CB"))
                                {
                                    // txtScan.Text = ValueText2(Gobal);
                                    nc.ProcessName = txtScan.Text;// ValueText2(Gobal);
                                }
                                else if (txtLineName.Text.Equals("TW02-SC_PB"))
                                {
                                    nc.ProcessName = txtScan.Text;// ValueText3(Gobal);
                                }
                                else
                                {
                                    // txtScan.Text = ValueText(Gobal);
                                    nc.ProcessName = txtScan.Text;// ValueText(Gobal);
                                }

                                nc.Seq = Gobal;
                                db.tb_QCCountPDs.InsertOnSubmit(nc);
                                db.SubmitChanges();
                            }
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
            try
            {
                if(e.RowIndex>0)
                {

                }
            }
            catch { }
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

        private void radButtonElement1_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("ต้องการล้างข้อมูลนี้ใหม่?", "ล้างข้อมูล", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    using (DataClasses1DataContext db = new DataClasses1DataContext())
                    {
                        var tbQ = db.tb_QCCountPDs.Where(p => p.WONo.Equals(txtWoNo.Text)).ToList();
                        foreach (var rd in tbQ)
                        {
                            tb_QCCountPD ps = db.tb_QCCountPDs.Where(p => p.id.Equals(rd.id)).FirstOrDefault();
                            if (ps != null)
                            {
                                db.tb_QCCountPDs.DeleteOnSubmit(ps);
                                db.SubmitChanges();
                            }
                        }
                    }
                }
            }
            catch { }
            LoadData();
        }

        private void radGridView4_CellEndEdit(object sender, Telerik.WinControls.UI.GridViewCellEventArgs e)
        {
            if(e.RowIndex>=0)
            {
                try
                {
                    using (DataClasses1DataContext db = new DataClasses1DataContext())
                    {
                        string DN = radGridView4.CurrentRow.Cells["DayN"].Value.ToString();
                        int Seq = 0;
                        int.TryParse(radGridView4.CurrentRow.Cells["Seq"].Value.ToString(), out Seq);
                        if (!DN.Equals(""))
                        {
                            if (DN.Equals("N") || DN.Equals("D"))
                            {
                                tb_QCCountPD ck = db.tb_QCCountPDs.Where(c => c.WONo.Equals(txtWoNo.Text) && c.Seq.Equals(Seq)).FirstOrDefault();
                                if(ck!=null)
                                {
                                    ck.DayN = DN;
                                    db.SubmitChanges();
                                }
                            }
                        }
                    }
                }
                catch { }
            }
        }
    }
}
