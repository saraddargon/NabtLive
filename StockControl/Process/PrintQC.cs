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
    public partial class PrintQC : Telerik.WinControls.UI.RadRibbonForm
    {
   
        Telerik.WinControls.UI.RadTextBox CodeNo_tt = new Telerik.WinControls.UI.RadTextBox();
        int screen = 0;
        public PrintQC(string WOx)
        {
            InitializeComponent();
            WO = WOx;
            screen = 1;
        }
        public PrintQC()
        {
            InitializeComponent();
        }

        string WO = "";
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
            //radDateTimePicker1.Value = DateTime.Now;
            //radDateTimePicker2.Value = DateTime.Now;
            radGridView1.AutoGenerateColumns = false;
            LoadData();
        }
        private void LoadData()
        {
            using (DataClasses1DataContext db = new DataClasses1DataContext())
            {
                radGridView1.DataSource = null;
                radGridView1.DataSource = db.sp_55_SelectQC(WO).ToList();
                
            }
        }
        private void btn_PrintPR_Click(object sender, EventArgs e)
        {
            try
            {
                if (radGridView1.CurrentRow.Cells["FromISO"].Value.ToString().Equals("FM-PD-026_1"))
                {
                    this.Cursor = Cursors.WaitCursor;
                    dbShowData.PrintData(radGridView1.CurrentRow.Cells["WONo"].Value.ToString()
                        , radGridView1.CurrentRow.Cells["PartNo"].Value.ToString()
                        , radGridView1.CurrentRow.Cells["QCNo"].Value.ToString());
                    this.Cursor = Cursors.Default;
                }
                else if (radGridView1.CurrentRow.Cells["FromISO"].Value.ToString().Equals("FM-PD-033_1"))
                {
                    this.Cursor = Cursors.WaitCursor;
                    dbShowData.PrintData033(radGridView1.CurrentRow.Cells["WONo"].Value.ToString()
                        , radGridView1.CurrentRow.Cells["PartNo"].Value.ToString()
                        , radGridView1.CurrentRow.Cells["QCNo"].Value.ToString());
                    this.Cursor = Cursors.Default;
                }
                else if (radGridView1.CurrentRow.Cells["FromISO"].Value.ToString().Equals("FM-PD-035_1"))
                {
                    this.Cursor = Cursors.WaitCursor;
                    dbShowData.PrintData035(radGridView1.CurrentRow.Cells["WONo"].Value.ToString()
                        , radGridView1.CurrentRow.Cells["PartNo"].Value.ToString()
                        , radGridView1.CurrentRow.Cells["QCNo"].Value.ToString());
                    this.Cursor = Cursors.Default;
                }
                else if (radGridView1.CurrentRow.Cells["FromISO"].Value.ToString().Equals("FM-QA-055_02_1"))
                {
                    this.Cursor = Cursors.WaitCursor;
                    dbShowData.PrintData5501(radGridView1.CurrentRow.Cells["WONo"].Value.ToString()
                        , radGridView1.CurrentRow.Cells["PartNo"].Value.ToString()
                        , radGridView1.CurrentRow.Cells["QCNo"].Value.ToString());
                    this.Cursor = Cursors.Default;
                }
                else if (radGridView1.CurrentRow.Cells["FromISO"].Value.ToString().Equals("FM-QA-056_02_1"))
                {
                    this.Cursor = Cursors.WaitCursor;
                    dbShowData.PrintData5601(radGridView1.CurrentRow.Cells["WONo"].Value.ToString()
                        , radGridView1.CurrentRow.Cells["PartNo"].Value.ToString()
                        , radGridView1.CurrentRow.Cells["QCNo"].Value.ToString());
                    this.Cursor = Cursors.Default;
                }
            }
            catch { this.Cursor = Cursors.Default; }
            this.Cursor = Cursors.Default;
        }

        private void btnExport_Click(object sender, EventArgs e)
        {

        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            LoadData();
        }
    }
}
