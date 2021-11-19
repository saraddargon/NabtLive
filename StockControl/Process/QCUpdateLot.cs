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
        public QCUpdateLot()
        {
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
            using (DataClasses1DataContext db = new DataClasses1DataContext())
            {
                tb_QCCheckMachine mc = db.tb_QCCheckMachines.Where(w => w.WONo.Equals(txtWoNo.Text) && w.Seq.Equals(48)).FirstOrDefault();
                if(mc!=null)
                {
                    txtLot.Text = Convert.ToString(mc.Value1);
                }
                tb_QCCheckMachine mc2 = db.tb_QCCheckMachines.Where(w => w.WONo.Equals(txtWoNo.Text) && w.Seq.Equals(49)).FirstOrDefault();
                if (mc2 != null)
                {
                    rdoWorkShift.Text = Convert.ToString(mc2.Value1);
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
                    db.sp_46_QCUpdateLot(txtWoNo.Text, txtLot.Text, rdoWorkShift.Text);
                    MessageBox.Show("บันทึกแล้ว");
                }
            }
        }
    }
}
