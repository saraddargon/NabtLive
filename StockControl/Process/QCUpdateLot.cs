using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Linq;
using Microsoft.VisualBasic.FileIO;
namespace StockControl
{
    public partial class QCUpdateLot : Telerik.WinControls.UI.RadRibbonForm
    {
   
        Telerik.WinControls.UI.RadTextBox CodeNo_tt = new Telerik.WinControls.UI.RadTextBox();
        int screen = 0;
        public QCUpdateLot(string  CodeNox)
        {
            InitializeComponent();            
            screen = 1;
            txtWoNo.Text = CodeNox;
        }
        public QCUpdateLot(string CodeNox,string FormISOx,string LotNox,string PartNox)
        {
            InitializeComponent();
            screen = 1;
            txtWoNo.Text = CodeNox;
            FormISO = FormISOx;
            LotNo = LotNox;
            PartNo = PartNox;
        }
        public QCUpdateLot()
        {
            InitializeComponent();
        }

        string PR1 = "";
        string PR2 = "";
        string Type = "";
        string FormISO = "";
        string LotNo = "";
        string PartNo = "";
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
            using (DataClasses1DataContext db = new DataClasses1DataContext())
            {
                if (FormISO.Equals("FM-PD-026_1"))
                {
                    txtQty.Enabled = false;
                    txtHight.Enabled = false;
                    tb_QCCheckMachine mc = db.tb_QCCheckMachines.Where(w => w.WONo.Equals(txtWoNo.Text) && w.Seq.Equals(48)).FirstOrDefault();
                    if (mc != null)
                    {
                        txtLot.Text = Convert.ToString(mc.Value1);
                    }
                    tb_QCCheckMachine mc2 = db.tb_QCCheckMachines.Where(w => w.WONo.Equals(txtWoNo.Text) && w.Seq.Equals(49)).FirstOrDefault();
                    if (mc2 != null)
                    {
                        rdoWorkShift.Text = Convert.ToString(mc2.Value1);
                    }
                }
                else
                {
                    rdoWorkShift.Enabled = false;
                    txtSetconner.Enabled = false;
                    txtLot.Text = LotNo;
                    string TypeReport = dbShowData.GetReportName("STD.Base", PartNo, FormISO);
                 
                    int IP1 = 35;
                    int IP2 = 36;               
                    //if (TypeReport.Equals("SPG"))
                    //{                       
                    //    IP1 = 35;
                    //    IP2 = 36;
                    //}

                    //tb_QCCheckMachine mc = db.tb_QCCheckMachines.Where(w => w.WONo.Equals(txtWoNo.Text) && w.Seq.Equals(48)).FirstOrDefault();
                    //if (mc != null)
                    //{
                    //    txtLot.Text = Convert.ToString(mc.Value1);
                    //}
                    tb_QCCheckMachine mc2 = db.tb_QCCheckMachines.Where(w => w.WONo.Equals(txtWoNo.Text) && w.Seq.Equals(IP1)).FirstOrDefault();
                    if (mc2 != null)
                    {
                        txtQty.Text = Convert.ToString(mc2.Value1);
                    }
                    tb_QCCheckMachine mc3 = db.tb_QCCheckMachines.Where(w => w.WONo.Equals(txtWoNo.Text) && w.Seq.Equals(IP2)).FirstOrDefault();
                    if (mc3 != null)
                    {
                        txtHight.Text = Convert.ToString(mc3.Value1);
                    }
                }


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
                    if (FormISO.Equals("FM-PD-026_1"))
                    {
                        db.sp_46_QCUpdateLot(txtWoNo.Text, txtLot.Text, rdoWorkShift.Text);
                        tb_QCCheckMachine chk = db.tb_QCCheckMachines.Where(p => p.WONo.Equals(txtWoNo.Text) && p.Seq.Equals(42)).FirstOrDefault();
                        if (chk != null)
                        {
                            chk.Value1 = txtSetconner.Text;
                            db.SubmitChanges();
                        }
                    }
                    else if (FormISO.Equals("FM-PD-001"))
                    {
                        string TypeReport = dbShowData.GetReportName("STD.Base", PartNo, FormISO);
                        int SQR = 45;
                        int IP1 = 35;
                        int IP2 = 36;
                        if (TypeReport.Equals("STD.PPC"))
                        {
                            SQR = 39;
                        }
                        if (TypeReport.Equals("SPG"))
                        {
                            SQR = 45;
                           
                        }

                        //34,35,41
                        tb_QCCheckMachine chk1 = db.tb_QCCheckMachines.Where(p => p.WONo.Equals(txtWoNo.Text) && p.Seq.Equals(IP1)).FirstOrDefault();
                        if (chk1 != null)
                        {
                            chk1.Value1 = txtQty.Text;
                            db.SubmitChanges();
                        }
                        tb_QCCheckMachine chk2 = db.tb_QCCheckMachines.Where(p => p.WONo.Equals(txtWoNo.Text) && p.Seq.Equals(IP2)).FirstOrDefault();
                        if (chk2 != null)
                        {
                            chk2.Value1 = txtHight.Text;
                            db.SubmitChanges();
                        }
                        
                        

                        tb_QCCheckMachine chk3 = db.tb_QCCheckMachines.Where(p => p.WONo.Equals(txtWoNo.Text) && p.Seq.Equals(SQR)).FirstOrDefault();
                        if (chk3 != null)
                        {
                            chk3.Value1 = txtLot.Text;
                            db.SubmitChanges();
                        }
                    }
                    MessageBox.Show("บันทึกแล้ว");
                }
            }
        }
    }
}
