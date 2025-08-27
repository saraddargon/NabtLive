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
using System.Runtime.InteropServices;
namespace StockControl
{
    public partial class QCFormPD026 : Telerik.WinControls.UI.RadRibbonForm
    {
        public QCFormPD026()
        {
            InitializeComponent();
        }
        public QCFormPD026(string Wox, string FormISOx, string PTAGx, string LineNamex,string Typex,string PTAGx2)
        {
            this.Text = "Check Sheet " + FormISOx;
            InitializeComponent();
            WOs = Wox;
            FormISO = FormISOx;
            PTAG = PTAGx;
            PTAG2 = PTAGx2;
            TypeP = Typex;
            LineName = LineNamex;
            OpenPage = 0;
        }
        string TypeP = "";
        string WOs = "";
        string FormISO = "";
        string LineName = "TW01-PB";
        string PTAG = "";
        string PTAG2 = "";
        string SPG33_1 = "８２００～９７２０　N";
        string SPG33_2 = "７０６０～８４２０　N";
        string Piggy = "Piggy Back Checksheet การตรวจสอบด้วยตนเอง　（Size 24）";
        string LotMark = "Lot ที่ตอกสามารถอ่านได้อย่างชัดเจน ";
        int OpenPage = 0;
     
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.S))
            {
                // Alt+F pressed
                //  ClearData();
                SaveData();
                //return false;
                //txtSeriesNo.Focus();
            }
            else if (keyData == (Keys.F8))
            {
                SelectNew();
            }
            else if ((keyData == (Keys.Control | Keys.N)))
            {
                SelectNew();
            }
            else if (keyData == (Keys.F9))
            {
                SaveData();
            }
            else if (keyData == (Keys.F5))
            {

            }
            else if (keyData == (Keys.F10))
            {

            }
            else if (keyData == (Keys.Escape))
            {
                this.Close();
            }
            else if ((keyData == (Keys.Control | Keys.Tab)))
            {
                string AciveC = radPageView1.SelectedPage.Name;

                if (AciveC.Equals("radPageViewPage1"))
                {
                    OpenPage = 1;
                    NextPage(radPageViewPage7);
                }                
                else
                {
                    OpenPage = 1;
                    NextPage(radPageViewPage1);
                }
            }
            else if (keyData == Keys.Up)
            {
                UpDownData(0);
            }
            else if (keyData == Keys.Down)
            {
                UpDownData(1);
            }

          
            return base.ProcessCmdKey(ref msg, keyData);
        }

        int gobalNo = 0;
        int gobalNo2 = 0;
        private void ClearGobalNo()
        {
            gobalNo2 = 0;
            gobalNo = 0;
            txtToppic.Text = "";
            txtSetData.Text = "";
            txtDataBox.Text = "";
            txtValue.Text = "";
            txtNGQty.Text = "";
            txtRemark.Text = "";
            txtRank.Text = "";
        }
        private void UpDownData(int Ac)
        {
            try
            {
                
                int ValueID = 0;
                int qid = 0;
                int Cseq = 0;
                int.TryParse(txtNGID.Text, out qid);
                int orderby = 0;
                int CCK = 0;
                string TypeRP = "";
                // decimal Value1 = 0;
                // decimal Value2 = 0;
                if(txtValueCheck.Text.Equals("Yes") && lblInvalidValue.Text.Equals("NG"))
                {
                    if (FormISO.Equals("FM-PD-156") || FormISO.Equals("FM-PD-157") || FormISO.Equals("FM-PD-033") || FormISO.Equals("FM-PD-033_1")
                        || FormISO.Equals("FM-PD-003") || FormISO.Equals("FM-PD-163") || FormISO.Equals("FM-PD-011"))
                    {
                        //PD-156 , PD-157, PD-033 , PD-003 , PD-163 , PD-011 
                        if(MessageBox.Show("ไม่สามารถไปข้อต่อไปได้ถ้าใส่ค่าไม่ถูกต้อง","Error",MessageBoxButtons.OK,MessageBoxIcon.Error)==DialogResult.OK)
                        {

                        }
                        return;
                    }
                }


                if (txtValueCheck.Text.Equals("Yes") && txtDataBox.Text.Equals("") && !txtToppic.Text.Equals(""))
                {
                    if (FormISO.Equals("FM-PD-035_1"))
                    {
                        using (DataClasses1DataContext db = new DataClasses1DataContext())
                        {
                            var tkg = db.tb_QCTAGs.Where(q => q.QCNo.Equals(txtQCNo.Text)).ToList();
                            if (tkg.Count <= 1)
                            {
                                CCK = 1;
                                MessageBox.Show("โปรดใส่ค่าในช่องก่อน!", "Please. Insert Data.", MessageBoxButtons.OK, MessageBoxIcon.Error);

                            }
                            else
                            {
                                //MessageBox.Show("โปรดใส่ค่าในช่องก่อน!", "Please. Insert Data.", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }

                        }
                    }
                    else if (FormISO.Equals("FM-PD-157"))
                    {
                        using (DataClasses1DataContext db = new DataClasses1DataContext())
                        {
                            //1 = 4567
                            //2,3 =467
                            //tb_QCGroupPart pds = db.tb_QCGroupParts.Where(p => p.FormISO.Equals(FormISO) && p.SetDate2.Equals("YES")).FirstOrDefault();
                            //if (pds != null)
                            //{
                            tb_QCTAG pds = db.tb_QCTAGs.Where(p => p.BarcodeTag.Equals(txtScanTAG.Text)).FirstOrDefault();
                            if (pds != null)
                            {
                                tb_QCGroupPart pds2 = db.tb_QCGroupParts.Where(p => p.FormISO.Equals(FormISO) && p.PartNo.Equals(txtPartNo.Text) &&
                                (p.SetData.Equals("SPG") || p.SetData.Equals("TOHO"))).FirstOrDefault();
                                if (pds2 != null)
                                {
                                    if (pds.GType.Equals("หัว"))
                                    {
                                        if (lblSeq.Text == "ลำดับ 4" || lblSeq.Text == "ลำดับ 6" || lblSeq.Text == "ลำดับ 7" || lblSeq.Text == "ลำดับ 5")
                                        {
                                            CCK = 1;
                                            MessageBox.Show("โปรดใส่ค่าในช่องก่อน!", "Please. Insert Data.", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        }
                                    }
                                    else
                                    {
                                        if (lblSeq.Text == "ลำดับ 5")
                                        {
                                            CCK = 1;
                                            MessageBox.Show("โปรดใส่ค่าในช่องก่อน!", "Please. Insert Data.", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        }
                                    }
                                }
                                else
                                {
                                    CCK = 1;
                                    MessageBox.Show("โปรดใส่ค่าในช่องก่อน!", "Please. Insert Data.", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }   
                                

                            }
                        }
                        // MessageBox.Show("โปรดใส่ค่าในช่องก่อน!", "Please. Insert Data.", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else if (FormISO.Equals("FM-PD-163"))
                    {
                        using (DataClasses1DataContext db = new DataClasses1DataContext())
                        {
                            if (txtOfTAG.Text.Equals("1of1") || txtOfTAG.Text.Contains("1of"))
                            {
                                CCK = 1;
                                MessageBox.Show("โปรดใส่ค่าในช่องก่อน!", "Please. Insert Data.", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            else
                            {
                                if (lblSeq.Text == "ลำดับ 1" || lblSeq.Text == "ลำดับ 2")
                                {
                                }
                                else
                                {
                                    CCK = 1;
                                    MessageBox.Show("โปรดใส่ค่าในช่องก่อน!", "Please. Insert Data.", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                        }
                    }
                    else if (FormISO.Equals("FM-PD-003"))
                    {
                        using (DataClasses1DataContext db = new DataClasses1DataContext())
                        {
                            //var tkg = db.tb_QCTAGs.Where(q => q.QCNo.Equals(txtQCNo.Text)).ToList();
                            if (txtOfTAG.Text.Equals("1of1") || txtOfTAG.Text.Contains("1of"))
                            {
                                CCK = 1;
                                MessageBox.Show("โปรดใส่ค่าในช่องก่อน!", "Please. Insert Data.", MessageBoxButtons.OK, MessageBoxIcon.Error);

                            }
                            else
                            {
                                if (lblSeq.Text == "ลำดับ 6" || lblSeq.Text == "ลำดับ 7" || lblSeq.Text == "ลำดับ 8")
                                {
                                    string STEP1 = cboCheckGroupPart.Text;
                                    tb_QCGroupPart ps = db.tb_QCGroupParts.Where(p => p.PartNo.Equals(txtPartNo.Text) && p.FormISO.Equals("FM-PD-003")
                                    && p.StepPart.Equals(STEP1)
                                    && p.SetDate2.Equals("Yes")
                                    ).FirstOrDefault();
                                    if (ps != null)
                                    {
                                        CCK = 1;
                                        MessageBox.Show("โปรดใส่ค่าในช่องก่อน!", "Please. Insert Data.", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    }
                                }
                            }

                        }
                    }
                    else if (FormISO.Equals("FM-PD-033_1"))
                    {
                        using (DataClasses1DataContext db = new DataClasses1DataContext())
                        {
                            //var tkg = db.tb_QCTAGs.Where(q => q.QCNo.Equals(txtQCNo.Text)).ToList();
                            //if (tkg.Count <= 1)
                            //{
                            //  //  var rs = tkg.FirstOrDefault();
                            //    CCK = 1;
                            //    MessageBox.Show("โปรดใส่ค่าในช่องก่อน!", "Please. Insert Data.", MessageBoxButtons.OK, MessageBoxIcon.Error);

                            //}
                            //else
                            //{
                            //    if (lblSeq.Text == "ลำดับ 1")
                            //    {
                            //        CCK = 1;
                            //        MessageBox.Show("โปรดใส่ค่าในช่องก่อน!", "Please. Insert Data.", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            //    }
                            //}
                            if (txtOfTAG.Text.Equals("No 1"))
                            {
                                MessageBox.Show("โปรดใส่ค่าในช่องก่อน!", "Please. Insert Data.", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            else
                            {
                                if (lblSeq.Text == "ลำดับ 1")
                                {
                                    CCK = 1;
                                    MessageBox.Show("โปรดใส่ค่าในช่องก่อน!", "Please. Insert Data.", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }

                        }
                    }
                    else
                    {
                        CCK = 1;
                        MessageBox.Show("โปรดใส่ค่าในช่องก่อน!", "Please. Insert Data.", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                }

                if (CCK == 0)
                {
                    txtRank.Text = "";
                    txtValue.Text = "";
                    txtNGQty.Text = "";
                    txtToppic.Text = "";
                    txtSetData.Text = "";
                    lblSeq.Text = "";
                    txtDataBox.Text = "";
                    txtRemark.Text = "";
                    lblInvalidValue.Text = "";
                    if (Ac == 0)
                    {
                        gobalNo -= 1;
                        if (gobalNo < 0)
                            gobalNo = 0;
                    }
                    else
                    {
                        gobalNo += 1;
                        if (gobalNo > (gobalNo2 + 1))
                        {
                            gobalNo -= 1;
                        }
                    }
                    if (gobalNo == 0)
                        gobalNo = 1;

                    using (DataClasses1DataContext db = new DataClasses1DataContext())
                    {
                        // var liqg = db.tb_QCGroupParts.Where(q => q.FormISO.Equals(FormISO) && q.PartNo.Equals(txtPartNo.Text)).OrderBy(o=>o.Seq).ToList();
                        var liqg = db.sp_46_QCMaster_Select(FormISO, txtPartNo.Text).ToList();
                        foreach (var rd in liqg)
                        {
                            ValueID = rd.id;
                            if (rd.id > 0)
                            {
                                Cseq += 1;
                                if (gobalNo == Cseq)
                                {
                                    orderby = ValueID;
                                }
                            }
                        }

                        //if(orderby>0)
                        //{                       
                        cboCheckGroupPart.SelectedValue = orderby;
                        //}

                    }
                }
                    
            }
            catch { }
        }
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
        string Path = "";
        private void Unit_Load(object sender, EventArgs e)
        {
            //เริ่มต้นให้ตรวจสอบว่ามีเอกสารถูกสร้างหรือยัง ถ้ายังให้สร้างขั้นมา ให้ใช้ Class DataClass (dbShowData.cs) เป็นตัวทำ
            try
            {
                ClearGobalNo();
                lblIns2.Visible = false;
                radGridView3.Visible = false;
                txtChecker.Visible = false;
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    tb_Path ph = db.tb_Paths.Where(p => p.PathCode.Equals("QCImage")).FirstOrDefault();
                    if(ph!=null)
                    {
                        Path = ph.PathFile;
                    }
                }

                    SetValueFormISO();
                dbShowData.CreateListQC(WOs, FormISO);
               

                txtTempNo.Text = dbClss.GetSeriesNo(89, 2);
                radPageViewPage1.Enabled = true;
                radPageViewPage7.Enabled = true;
                groupBox1.Text = "Detail Production Order -> " + FormISO;
                if (!PTAG.Equals("New"))
                {
                    LoadData();
                   // LoadChecker();
                    SetValueFormISO2();
                    radButton1_Click_1(sender, e);
                    txtInspector.Focus();
                    LotMark = LotMark + " ( " + txtLotNo.Text + " )";
                }
                else
                {
                    txtScanTAG.Text = "";
                    txtScanTAG.Focus();
                }
                ShowDefaultPage();


            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }
        private void ShowDefaultPage()
        {
            if (FormISO.Equals("FM-PD-035_1")
                || FormISO.Equals("FM-PD-002")
                 || FormISO.Equals("FM-PD-112")
                  || FormISO.Equals("FM-PD-113")
                  || FormISO.Equals("FM-PD-123")
                  || FormISO.Equals("FM-PD-153")
                  || FormISO.Equals("FM-PD-010")
                  || FormISO.Equals("FM-PD-164")
                )
            {
                this.radPageView1.SelectedPage=radPageViewPage7;
                Snew += 1;
                OpenPage = 1;
                txtDataBox.Focus();
            }
        }
        private void SetValueFormISO()
        {
            try
            {
                cboSelectCheckBy.Items.Clear();
               // cboSelectCheckBy.Items.Add("");
                cboSelectCheckBy.Items.Add("Inspector");
                cboSelectCheckBy.Items.Add("Check By");
                cboSelectCheckBy.Text = "Inspector";
                // ribbon026.Visibility = Telerik.WinControls.ElementVisibility.Collapsed;
                if (FormISO.Equals("FM-PD-026_1"))
                {

                    cboSelectCheckBy.Items.Clear();
                    // cboSelectCheckBy.Items.Add("");
                    cboSelectCheckBy.Items.Add("ผู้จัดทำเอกสาร");
                    cboSelectCheckBy.Items.Add("ผู้ตรวจสอบก่อนผลิต");
                    cboSelectCheckBy.Items.Add("พนักงานประกอบ SUB LINE");
                    cboSelectCheckBy.Items.Add("พนักงานประกอบ MAIN LINE");
                    cboSelectCheckBy.Items.Add("พนักงานประกอบ FINAL LINE");
                    cboSelectCheckBy.Text = "ผู้จัดทำเอกสาร";

                }
                else if (FormISO.Equals("FM-PD-033_1"))
                {
                    cboSelectCheckBy.Items.Clear();
                    // cboSelectCheckBy.Items.Add("");
                    cboSelectCheckBy.Items.Add("ผู้ตรวจสอบ หัว");
                    cboSelectCheckBy.Items.Add("ผู้ตรวจสอบ กลาง");
                    cboSelectCheckBy.Items.Add("ผู้ตรวจสอบ ท้าย");
                    cboSelectCheckBy.Text = "ผู้ตรวจสอบ หัว";

                }
                else if (FormISO.Equals("FM-PD-035_1"))
                {
                    cboSelectCheckBy.Items.Clear();
                    // cboSelectCheckBy.Items.Add("");
                    cboSelectCheckBy.Items.Add("ผู้ตรวจสอบ");
                    cboSelectCheckBy.Items.Add("พนักงานตรวจ ก่อนผลิต");
                    cboSelectCheckBy.Items.Add("พนักงานตรวจ หลังผลิต");
                    cboSelectCheckBy.Text = "ผู้ตรวจสอบ";
                }
                else if (FormISO.Equals("FM-PD-001"))
                {
                    cboSelectCheckBy.Items.Clear();
                    // cboSelectCheckBy.Items.Add("");
                    cboSelectCheckBy.Items.Add("ผู้จัดทำเอกสาร");
                    cboSelectCheckBy.Items.Add("ผู้ตรวจสอบก่อนผลิต");

                    //cboSelectCheckBy.Items.Add("ขัน Plug");
                    //cboSelectCheckBy.Items.Add("Sub Line");
                    cboSelectCheckBy.Items.Add("Stamp Lot /ประกอบ");
                    cboSelectCheckBy.Items.Add("Test Leak");
                    cboSelectCheckBy.Items.Add("ตรวจสอบท้ายไลน์");
               

                    cboSelectCheckBy.Items.Add("พนักงานประกอบ SUB LINE");
                    cboSelectCheckBy.Items.Add("พนักงานตรวจสอบ SUB LINE 1");
                    cboSelectCheckBy.Items.Add("พนักงานตรวจสอบ SUB LINE 2");
                    cboSelectCheckBy.Items.Add("พนักงานขัน Plug");

                   // cboSelectCheckBy.Items.Add("สลับเบรค");
                    cboSelectCheckBy.Text = "ผู้จัดทำเอกสาร";
                }
                else if (FormISO.Equals("FM-PD-002"))
                {
                    cboSelectCheckBy.Items.Clear();
                    // cboSelectCheckBy.Items.Add("");
                    cboSelectCheckBy.Text = "Check By";
                    cboSelectCheckBy.Items.Add("Check By");
                    cboSelectCheckBy.Items.Add("ผู้ตรวจสอบ 1");
                    cboSelectCheckBy.Items.Add("ผู้ตรวจสอบ 2");
                    cboSelectCheckBy.Items.Add("ผู้ตรวจสอบ 3");
                    cboSelectCheckBy.Items.Add("ผู้ตรวจสอบ 4");
                    cboSelectCheckBy.Items.Add("ผู้ตรวจสอบ 5");
                }
                else if (FormISO.Equals("FM-PD-003"))
                {
                    cboSelectCheckBy.Items.Clear();
                    // cboSelectCheckBy.Items.Add("");
                    cboSelectCheckBy.Text = "ผู้ตรวจสอบ 1";
                    cboSelectCheckBy.Items.Add("ผู้ตรวจสอบ 1");
                    cboSelectCheckBy.Items.Add("ผู้ตรวจสอบ 2");
                }
                else if (FormISO.Equals("FM-PD-003_S") || FormISO.Equals("FM-PD-156") || FormISO.Equals("FM-PD-011")
                    || FormISO.Equals("FM-PD-157")
                    || FormISO.Equals("FM-PD-163")
                    || FormISO.Equals("FM-PD-077")
                    || FormISO.Equals("FM-PD-139"))
                {
                    cboSelectCheckBy.Items.Clear();
                    // cboSelectCheckBy.Items.Add("");
                  
                    cboSelectCheckBy.Items.Add("ผู้ตรวจสอบ 1");
                    cboSelectCheckBy.Items.Add("ผู้ตรวจสอบ 2");


                    if (FormISO.Equals("FM-PD-077"))
                    {
                        cboSelectCheckBy.Items.Add("อนุมัติการปล่อยงาน");
                    }
                    if (FormISO.Equals("FM-PD-139"))
                    {
                        cboSelectCheckBy.Items.Add("ผู้ตรวจสอบ 3");
                    }


                    cboSelectCheckBy.Text = "ผู้ตรวจสอบ 1";
                }
                else if (FormISO.Equals("FM-PD-109"))
                {

                    cboSelectCheckBy.Items.Clear();
                    // cboSelectCheckBy.Items.Add("");
                    cboSelectCheckBy.Items.Add("ผู้จัดทำเอกสาร");
                    cboSelectCheckBy.Items.Add("ผู้ตรวจสอบก่อนผลิต");
                    cboSelectCheckBy.Items.Add("Sub line");
                    cboSelectCheckBy.Items.Add("Main line");
                    cboSelectCheckBy.Items.Add("ตรวจสอบท้ายไลน์");
                    cboSelectCheckBy.Text = "ผู้จัดทำเอกสาร";


                }
                else if (FormISO.Equals("FM-PD-110"))
                {

                    cboSelectCheckBy.Items.Clear();
                    // cboSelectCheckBy.Items.Add("");
                    cboSelectCheckBy.Items.Add("ผู้จัดทำเอกสาร");
                    cboSelectCheckBy.Items.Add("ผู้ตรวจสอบก่อนผลิต");
                    cboSelectCheckBy.Items.Add("ประกอบ");
                    cboSelectCheckBy.Items.Add("Test Leak");
                    cboSelectCheckBy.Items.Add("ตรวจสอบท้ายไลน์");
                    cboSelectCheckBy.Text = "ผู้จัดทำเอกสาร";


                }
                else if (FormISO.Equals("FM-PD-112")||FormISO.Equals("FM-PD-113") || FormISO.Equals("FM-PD-123") || FormISO.Equals("FM-PD-153") || FormISO.Equals("FM-PD-010")
                    || FormISO.Equals("FM-PD-164"))
                {
                    cboSelectCheckBy.Items.Clear();
                    // cboSelectCheckBy.Items.Add("");                  
                    cboSelectCheckBy.Items.Add("ผู้ตรวจสอบ");
                    cboSelectCheckBy.Items.Add("Check");
                    cboSelectCheckBy.Text = "ผู้ตรวจสอบ";
                }
                else if (FormISO.Equals("FM-PD-096"))
                {
                    cboSelectCheckBy.Items.Clear();
                    // cboSelectCheckBy.Items.Add("");
                    cboSelectCheckBy.Items.Add("ผู้จัดทำเอกสาร");
                    cboSelectCheckBy.Items.Add("ผู้ตรวจสอบก่อนผลิต");
                   // cboSelectCheckBy.Items.Add("PS1-PS3"); 
                    cboSelectCheckBy.Text = "ผู้จัดทำเอกสาร";

                }
                else if (FormISO.Equals("FM-PD-095"))
                {

                    cboSelectCheckBy.Items.Clear();
                    // cboSelectCheckBy.Items.Add("");
                    cboSelectCheckBy.Items.Add("ผู้จัดทำเอกสาร");
                    cboSelectCheckBy.Items.Add("ผู้ตรวจสอบก่อนผลิต");
                   // cboSelectCheckBy.Items.Add("ผู้ตรวจสอบ");
                    cboSelectCheckBy.Text = "ผู้จัดทำเอกสาร";


                }
                else if (FormISO.Equals("FM-PD-122"))
                {

                    cboSelectCheckBy.Items.Clear();
                    // cboSelectCheckBy.Items.Add("");
                    cboSelectCheckBy.Items.Add("ผู้จัดทำเอกสาร");
                    cboSelectCheckBy.Items.Add("ผู้ตรวจสอบก่อนผลิต");
                    cboSelectCheckBy.Items.Add("ประกอบ");
                    cboSelectCheckBy.Items.Add("ท้ายไลน์");
                    // cboSelectCheckBy.Items.Add("ผู้ตรวจสอบ");
                    cboSelectCheckBy.Text = "ผู้จัดทำเอกสาร";


                }
                else if (FormISO.Equals("FM-PD-013"))
                {

                    cboSelectCheckBy.Items.Clear();
                    // cboSelectCheckBy.Items.Add("");
                    cboSelectCheckBy.Items.Add("ผู้จัดเตรียม Part");
                    cboSelectCheckBy.Items.Add("ผู้ตรวจสอบก่อนผลิต");
                    cboSelectCheckBy.Items.Add("ผู้จัดทำเอกสาร");
                    cboSelectCheckBy.Items.Add("ประกอบ 1");
                    cboSelectCheckBy.Items.Add("ประกอบ 2");
                    cboSelectCheckBy.Items.Add("Test Leak");
                    cboSelectCheckBy.Items.Add("ท้ายไลน์");
                    cboSelectCheckBy.Text = "ผู้จัดทำเอกสาร";


                }
                else if (FormISO.Equals("FM-PD-014"))
                {

                    cboSelectCheckBy.Items.Clear();
                    // cboSelectCheckBy.Items.Add("");
                    cboSelectCheckBy.Items.Add("ผู้จัดเตรียม Part");
                    cboSelectCheckBy.Items.Add("ผู้จัดทำเอกสาร");
                    cboSelectCheckBy.Items.Add("ผู้ตรวจสอบก่อนผลิต");
                    cboSelectCheckBy.Items.Add("ประกอบ");                   
                    cboSelectCheckBy.Text = "ผู้จัดทำเอกสาร";


                }
                else if (FormISO.Equals("FM-PD-140"))
                {

                    cboSelectCheckBy.Items.Clear();
                    // cboSelectCheckBy.Items.Add("");
                    cboSelectCheckBy.Items.Add("ผู้จัดเตรียม Part");
                    cboSelectCheckBy.Items.Add("ผู้จัดทำเอกสาร");
                    cboSelectCheckBy.Items.Add("ผู้ตรวจสอบก่อนผลิต");
                    cboSelectCheckBy.Items.Add("ประกอบ");
                    cboSelectCheckBy.Text = "ผู้จัดทำเอกสาร";


                }
            }
            catch { }
        }
        private void SetValueFormISO2()
        {
            try
            {
                if (PTAG2 == "")
                    PTAG2 = PTAG;

                if (FormISO.Equals("FM-PD-026_1"))
                {
                    TypeP = "PD";
                    string TAG1 = "PQC," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",1of1," + txtPartNo.Text.ToUpper() + ",026_1";
                    dbShowData.InsertTAG(TAG1, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, "1of1", "PD", LineName, "Machine", "None");
                }
                else if(FormISO.Equals("FM-PD-001"))
                {
                    TypeP = "PD";
                    string TAG1 = "PQC," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",1of1," + txtPartNo.Text.ToUpper() + ",FMPD001";
                    dbShowData.InsertTAG(TAG1, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, "1of1", "PD", LineName, "Machine", "None");
                }
                else if (FormISO.Equals("FM-PD-109"))
                {
                    TypeP = "PD";
                    string TAG1 = "PQC," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",1of1," + txtPartNo.Text.ToUpper() + ",FMPD109";
                    dbShowData.InsertTAG(TAG1, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, "1of1", "PD", LineName, "Machine", "None");
                }
                else if (FormISO.Equals("FM-PD-110"))
                {
                    TypeP = "PD";
                    string TAG1 = "PQC," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",1of1," + txtPartNo.Text.ToUpper() + ",FMPD110";
                    dbShowData.InsertTAG(TAG1, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, "1of1", "PD", LineName, "Machine", "None");
                }
                else if (FormISO.Equals("FM-PD-013"))
                {
                    TypeP = "PD";
                    string TAG1 = "PQC," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",1of1," + txtPartNo.Text.ToUpper() + ",FMPD013";
                    dbShowData.InsertTAG(TAG1, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, "1of1", "PD", LineName, "Machine", "None");
                }
                else if (FormISO.Equals("FM-PD-122"))
                {
                    TypeP = "PD";
                    string TAG1 = "PQC," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",1of1," + txtPartNo.Text.ToUpper() + ",FMPD122";
                    dbShowData.InsertTAG(TAG1, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, "1of1", "PD", LineName, "Machine", "None");
                }
                else if (FormISO.Equals("FM-PD-014"))
                {
                    TypeP = "PD";
                    string TAG1 = "PQC," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",1of1," + txtPartNo.Text.ToUpper() + ",FMPD014";
                    dbShowData.InsertTAG(TAG1, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, "1of1", "PD", LineName, "Machine", "None");
                }
                else if (FormISO.Equals("FM-PD-096") || FormISO.Equals("FM-PD-095"))
                {
                    TypeP = "PD";
                    string TAG1 = "PQC," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",1of1," + txtPartNo.Text.ToUpper() + ",FMPD0956";
                    dbShowData.InsertTAG(TAG1, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, "1of1", "PD", LineName, "Machine", "None");
                }
                if (FormISO.Equals("FM-PD-140"))
                {
                    TypeP = "PD";
                    string TAG1 = "PQC," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",1of1," + txtPartNo.Text.ToUpper() + ",FMPD140";
                    dbShowData.InsertTAG(TAG1, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, "1of1", "PD", LineName, "Machine", "None");
                }
                else if (FormISO.Equals("FM-PD-033_1"))
                {
                    double HQty = 0;
                    int midle = 2;
                    double LotSize = 0;
                    double a = 0;
                    double ap = 0;
                    double.TryParse(txtQty.Text, out HQty);
                    double.TryParse(txtLotSize.Text, out LotSize);
                    a = 0;
                    ap = (Convert.ToDouble(HQty) % 2);
                    if (ap > 0)
                        a = 1;
                    midle = Convert.ToInt32(Math.Floor((Convert.ToDouble(HQty) / 2)) + a);//.ToString("###");
                    string Tof1 = "No 1";
                    string Tof2 = "No "+midle.ToString("###");
                    string Tof3 = "No "+HQty.ToString("###");

                    string TAG1 = "Head," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",No 1," + txtPartNo.Text.ToUpper() + ",033_1";
                    string TAG2 = "Middle," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",No " + midle.ToString("###") + "," + txtPartNo.Text.ToUpper() + ",033_1";
                    string TAG3 = "End," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",No " + HQty.ToString("###") + "," + txtPartNo.Text.ToUpper() + ",033_1";

                    dbShowData.InsertTAG(TAG1, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, Tof1, "PD", LineName, "หัว", PTAG2);
                    dbShowData.InsertTAG(TAG2, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, Tof2, "PD", LineName, "กลาง", PTAG2);
                    dbShowData.InsertTAG(TAG3, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, Tof3, "PD", LineName, "ท้าย", PTAG2);

                }
                else if (FormISO.Equals("FM-PD-035_1"))
                {
                    TypeP = "PD";
                    decimal QtyT = 0;
                    decimal.TryParse(txtQtyofTAG.Text, out QtyT);
                    if (QtyT == 0)
                        QtyT = 1;
                    if (!txtOfTAG.Text.Equals(""))
                    {
                        string[] DATAG = PTAG.Split(',');
                        if (DATAG.Length == 8)
                        {
                            if(!DATAG[0].Equals("PQC"))
                                dbShowData.InsertTAG(PTAG, txtProdNo.Text.ToUpper(), txtQCNo.Text, QtyT, txtOfTAG.Text, "PD", LineName, "PDTAG", PTAG2);
                        }
                    }
                }
                else if (FormISO.Equals("FM-QA-055_02_1") || FormISO.Equals("FM-QA-055"))
                {
                    TypeP = "QC";
                    decimal QtyT = 0;
                    decimal.TryParse(txtQtyofTAG.Text, out QtyT);
                    if (QtyT == 0)
                        QtyT = 1;
                    if (!txtOfTAG.Text.Equals(""))
                    {
                        dbShowData.InsertTAG(PTAG, txtProdNo.Text.ToUpper(), txtQCNo.Text, QtyT, txtOfTAG.Text, "QC", LineName, "QCTAG", PTAG2);
                    }
                }
                else if (FormISO.Equals("FM-QA-056_02_1"))
                {
                    TypeP = "QC";
                    string TAG1 = "No1," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",1of5," + txtPartNo.Text.ToUpper() + ",056_1";
                    string TAG2 = "No2," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",2of5," + txtPartNo.Text.ToUpper() + ",056_1";
                    string TAG3 = "No3," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",3of5," + txtPartNo.Text.ToUpper() + ",056_1";
                    string TAG4 = "No4," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",4of5," + txtPartNo.Text.ToUpper() + ",056_1";
                    string TAG5 = "No5," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",5of5," + txtPartNo.Text.ToUpper() + ",056_1";

                    dbShowData.InsertTAG(TAG1, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, "1of5", "QC", LineName, "ตัวที่ 1", PTAG2);
                    dbShowData.InsertTAG(TAG2, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, "2of5", "QC", LineName, "ตัวที่ 2", PTAG2);
                    dbShowData.InsertTAG(TAG3, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, "3of5", "QC", LineName, "ตัวที่ 3", PTAG2);
                    dbShowData.InsertTAG(TAG4, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, "4of5", "QC", LineName, "ตัวที่ 4", PTAG2);
                    dbShowData.InsertTAG(TAG5, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, "5of5", "QC", LineName, "ตัวที่ 5", PTAG2);


                }
                else if (FormISO.Equals("FM-QA-091") 
                    || FormISO.Equals("FM-QA-092")
                    || FormISO.Equals("FM-QA-143")
                    || FormISO.Equals("FM-QA-144")
                    || FormISO.Equals("FM-QA-161")
                    )
                {
                    TypeP = "QC";
                    string TAG1 = "One," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",1of5," + txtPartNo.Text.ToUpper() + ",0QA_1";
                    string TAG2 = "Two," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",2of5," + txtPartNo.Text.ToUpper() + ",0QA_2";
                    string TAG3 = "Three," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",3of5," + txtPartNo.Text.ToUpper() + ",0QA_3";
                    string TAG4 = "Four," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",4of5," + txtPartNo.Text.ToUpper() + ",0QA_4";
                    string TAG5 = "Five," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",5of5," + txtPartNo.Text.ToUpper() + ",0QA_5";

                    dbShowData.InsertTAG(TAG1, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, "1of5", "QC", LineName, "ตัวที่ 1", PTAG2);
                    dbShowData.InsertTAG(TAG2, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, "2of5", "QC", LineName, "ตัวที่ 2", PTAG2);
                    dbShowData.InsertTAG(TAG3, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, "3of5", "QC", LineName, "ตัวที่ 3", PTAG2);
                    dbShowData.InsertTAG(TAG4, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, "4of5", "QC", LineName, "ตัวที่ 4", PTAG2);
                    dbShowData.InsertTAG(TAG5, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, "5of5", "QC", LineName, "ตัวที่ 5", PTAG2);


                }
                else if (FormISO.Equals("FM-PD-003"))
                {
                    double HQty = 0;
                    int midle = 2;
                    double LotSize = 0;
                    double a = 0;
                    double ap = 0;
                    double.TryParse(txtQty.Text, out HQty);
                    double.TryParse(txtLotSize.Text, out LotSize);
                    a = 0;
                    ap = (Convert.ToDouble(HQty) % 2);
                    if (ap > 0)
                        a = 1;
                    midle = Convert.ToInt32(Math.Floor((Convert.ToDouble(HQty) / 2)) + a);//.ToString("###");
                    string Tof1 = "No 1";
                    string Tof2 = "No " + midle.ToString("###");
                    string Tof3 = "No " + HQty.ToString("###");

                    string TAG1 = "Head," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",1of1," + txtPartNo.Text.ToUpper() + ",FMPD003";
                    string TAG2 = "Middle," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",2of" + midle.ToString("###") + "," + txtPartNo.Text.ToUpper() + ",FMPD003";
                    string TAG3 = "End," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",3of" + HQty.ToString("###") + "," + txtPartNo.Text.ToUpper() + ",FMPD003";

                    dbShowData.InsertTAG(TAG1, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, Tof1, "PD", LineName, "หัว", PTAG2);
                    dbShowData.InsertTAG(TAG2, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, Tof2, "PD", LineName, "กลาง", PTAG2);
                    dbShowData.InsertTAG(TAG3, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, Tof3, "PD", LineName, "ท้าย", PTAG2);

                }
                else if (FormISO.Equals("FM-PD-003_S"))
                {
                    double HQty = 0;
                    int midle = 2;
                    double LotSize = 0;
                    double a = 0;
                    double ap = 0;
                    double.TryParse(txtQty.Text, out HQty);
                    double.TryParse(txtLotSize.Text, out LotSize);
                    a = 0;
                    ap = (Convert.ToDouble(HQty) % 2);
                    if (ap > 0)
                        a = 1;
                    midle = Convert.ToInt32(Math.Floor((Convert.ToDouble(HQty) / 2)) + a);//.ToString("###");
                    string Tof1 = "No 1";
                    string Tof2 = "No " + midle.ToString("###");
                    string Tof3 = "No " + HQty.ToString("###");

                    string TAG1 = "Head," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",1of1," + txtPartNo.Text.ToUpper() + ",FMPD003_S";
                    string TAG2 = "Middle," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",2of" + midle.ToString("###") + "," + txtPartNo.Text.ToUpper() + ",FMPD003_S";
                    string TAG3 = "End," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",3of" + HQty.ToString("###") + "," + txtPartNo.Text.ToUpper() + ",FMPD003_S";

                    dbShowData.InsertTAG(TAG1, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, Tof1, "PD", LineName, "หัว", TAG1);
                    dbShowData.InsertTAG(TAG2, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, Tof2, "PD", LineName, "กลาง", TAG2);
                    dbShowData.InsertTAG(TAG3, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, Tof3, "PD", LineName, "ท้าย", TAG3);

                }
                else if  (FormISO.Equals("FM-PD-156"))
                {
                    double HQty = 0;
                    int midle = 2;
                    double LotSize = 0;
                    double a = 0;
                    double ap = 0;
                    double.TryParse(txtQty.Text, out HQty);
                    double.TryParse(txtLotSize.Text, out LotSize);
                    a = 0;
                    ap = (Convert.ToDouble(HQty) % 2);
                    if (ap > 0)
                        a = 1;
                    midle = Convert.ToInt32(Math.Floor((Convert.ToDouble(HQty) / 2)) + a);//.ToString("###");
                    string Tof1 = "No 1";
                    string Tof2 = "No " + midle.ToString("###");
                    string Tof3 = "No " + HQty.ToString("###");

                    string TAG1 = "Head," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",1of1," + txtPartNo.Text.ToUpper() + ",FMPD156";
                    string TAG2 = "Middle," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",2of" + midle.ToString("###") + "," + txtPartNo.Text.ToUpper() + ",FMPD156";
                    string TAG3 = "End," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",3of" + HQty.ToString("###") + "," + txtPartNo.Text.ToUpper() + ",FMPD156";

                    dbShowData.InsertTAG(TAG1, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, Tof1, "PD", LineName, "หัว", TAG1);
                    dbShowData.InsertTAG(TAG2, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, Tof2, "PD", LineName, "กลาง", TAG2);
                    dbShowData.InsertTAG(TAG3, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, Tof3, "PD", LineName, "ท้าย", TAG3);
                }
                else if (FormISO.Equals("FM-PD-139"))
                {
                    double HQty = 0;
                    int midle = 2;
                    double LotSize = 0;
                    double a = 0;
                    double ap = 0;
                    double.TryParse(txtQty.Text, out HQty);
                    double.TryParse(txtLotSize.Text, out LotSize);
                    a = 0;
                    ap = (Convert.ToDouble(HQty) % 2);
                    if (ap > 0)
                        a = 1;
                    midle = Convert.ToInt32(Math.Floor((Convert.ToDouble(HQty) / 2)) + a);//.ToString("###");
                    string Tof1 = "No 1";
                    string Tof2 = "No " + midle.ToString("###");
                    string Tof3 = "No " + HQty.ToString("###");

                    string TAG1 = "Head," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",1of1," + txtPartNo.Text.ToUpper() + ",FMPD139";
                    string TAG2 = "Middle," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",2of" + midle.ToString("###") + "," + txtPartNo.Text.ToUpper() + ",FMPD139";
                    string TAG3 = "End," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",3of" + HQty.ToString("###") + "," + txtPartNo.Text.ToUpper() + ",FMPD139";

                    dbShowData.InsertTAG(TAG1, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, Tof1, "PD", LineName, "หัว", TAG1);
                    dbShowData.InsertTAG(TAG2, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, Tof2, "PD", LineName, "กลาง", TAG2);
                    dbShowData.InsertTAG(TAG3, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, Tof3, "PD", LineName, "ท้าย", TAG3);
                }
                else if (FormISO.Equals("FM-PD-011"))
                {
                    double HQty = 0;
                    int midle = 2;
                    double LotSize = 0;
                    double a = 0;
                    double ap = 0;
                    double.TryParse(txtQty.Text, out HQty);
                    double.TryParse(txtLotSize.Text, out LotSize);
                    a = 0;
                    ap = (Convert.ToDouble(HQty) % 2);
                    if (ap > 0)
                        a = 1;
                    midle = Convert.ToInt32(Math.Floor((Convert.ToDouble(HQty) / 2)) + a);//.ToString("###");
                    string Tof1 = "No 1";
                    string Tof2 = "No " + midle.ToString("###");
                    string Tof3 = "No " + HQty.ToString("###");

                    string TAG1 = "Head," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",1of1," + txtPartNo.Text.ToUpper() + ",FMPD011";
                    string TAG2 = "Middle," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",2of" + midle.ToString("###") + "," + txtPartNo.Text.ToUpper() + ",FMPD011";
                    string TAG3 = "End," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",3of" + HQty.ToString("###") + "," + txtPartNo.Text.ToUpper() + ",FMPD011";

                    dbShowData.InsertTAG(TAG1, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, Tof1, "PD", LineName, "หัว", TAG1);
                    dbShowData.InsertTAG(TAG2, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, Tof2, "PD", LineName, "กลาง", TAG2);
                    dbShowData.InsertTAG(TAG3, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, Tof3, "PD", LineName, "ท้าย", TAG3);
                }

                else if ((FormISO.Equals("FM-PD-157")))
                {
                    double HQty = 0;
                    int midle = 2;
                    double LotSize = 0;
                    double a = 0;
                    double ap = 0;
                    double.TryParse(txtQty.Text, out HQty);
                    double.TryParse(txtLotSize.Text, out LotSize);
                    a = 0;
                    ap = (Convert.ToDouble(HQty) % 2);
                    if (ap > 0)
                        a = 1;
                    midle = Convert.ToInt32(Math.Floor((Convert.ToDouble(HQty) / 2)) + a);//.ToString("###");
                    string Tof1 = "No 1";
                    string Tof2 = "No " + midle.ToString("###");
                    string Tof3 = "No " + HQty.ToString("###");

                    string TAG1 = "Head," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",1of1," + txtPartNo.Text.ToUpper() + ",FMPD157";
                    string TAG2 = "Middle," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",2of" + midle.ToString("###") + "," + txtPartNo.Text.ToUpper() + ",FMPD157";
                    string TAG3 = "End," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",3of" + HQty.ToString("###") + "," + txtPartNo.Text.ToUpper() + ",FMPD157";

                    dbShowData.InsertTAG(TAG1, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, Tof1, "PD", LineName, "หัว", PTAG2);
                    dbShowData.InsertTAG(TAG2, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, Tof2, "PD", LineName, "กลาง", PTAG2);
                    dbShowData.InsertTAG(TAG3, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, Tof3, "PD", LineName, "ท้าย", PTAG2);

                }
                else if ((FormISO.Equals("FM-PD-163")))
                {
                    double HQty = 0;
                    int midle = 2;
                    double LotSize = 0;
                    double a = 0;
                    double ap = 0;
                    double.TryParse(txtQty.Text, out HQty);
                    double.TryParse(txtLotSize.Text, out LotSize);
                    a = 0;
                    ap = (Convert.ToDouble(HQty) % 2);
                    if (ap > 0)
                        a = 1;
                    midle = Convert.ToInt32(Math.Floor((Convert.ToDouble(HQty) / 2)) + a);//.ToString("###");
                    string Tof1 = "No 1";
                    string Tof2 = "No " + midle.ToString("###");
                    string Tof3 = "No " + HQty.ToString("###");

                    string TAG1 = "Head," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",1of1," + txtPartNo.Text.ToUpper() + ",FMPD163";
                    string TAG2 = "Middle," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",2of" + midle.ToString("###") + "," + txtPartNo.Text.ToUpper() + ",FMPD163";
                    string TAG3 = "End," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",3of" + HQty.ToString("###") + "," + txtPartNo.Text.ToUpper() + ",FMPD163";

                    dbShowData.InsertTAG(TAG1, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, Tof1, "PD", LineName, "หัว", PTAG2);
                    dbShowData.InsertTAG(TAG2, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, Tof2, "PD", LineName, "กลาง", PTAG2);
                    dbShowData.InsertTAG(TAG3, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, Tof3, "PD", LineName, "ท้าย", PTAG2);

                }
                else if ((FormISO.Equals("FM-PD-077")))
                {
                    double HQty = 0;
                    int midle = 2;
                    double LotSize = 0;
                    double a = 0;
                    double ap = 0;
                    double.TryParse(txtQty.Text, out HQty);
                    double.TryParse(txtLotSize.Text, out LotSize);
                    a = 0;
                    ap = (Convert.ToDouble(HQty) % 2);
                    if (ap > 0)
                        a = 1;
                    midle = Convert.ToInt32(Math.Floor((Convert.ToDouble(HQty) / 2)) + a);//.ToString("###");
                    string Tof1 = "No 1";
                    string Tof2 = "No " + midle.ToString("###");
                    string Tof3 = "No " + HQty.ToString("###");

                    string TAG1 = "Head," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",1of1," + txtPartNo.Text.ToUpper() + ",FMPD077";
                   // string TAG2 = "Middle," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",2of" + midle.ToString("###") + "," + txtPartNo.Text.ToUpper() + ",FMPD157";
                    string TAG3 = "End," + txtProdNo.Text.ToUpper() + ",1,1," + txtLotNo.Text + ",2of" + HQty.ToString("###") + "," + txtPartNo.Text.ToUpper() + ",FMPD077";

                    dbShowData.InsertTAG(TAG1, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, Tof1, "PD", LineName, "หัว", PTAG2);
                  //  dbShowData.InsertTAG(TAG2, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, Tof2, "PD", LineName, "กลาง", PTAG2);
                    dbShowData.InsertTAG(TAG3, txtProdNo.Text.ToUpper(), txtQCNo.Text, 1, Tof2, "PD", LineName, "ท้าย", PTAG2);

                }
                else if (FormISO.Equals("FM-PD-002"))
                {
                    TypeP = "PD";
                    decimal QtyT = 0;
                    decimal.TryParse(txtQtyofTAG.Text, out QtyT);
                    if (QtyT == 0)
                        QtyT = 1;
                    if (!txtOfTAG.Text.Equals(""))
                    {
                        string[] DATAG = PTAG.Split(',');
                        if (DATAG.Length == 8)
                        {
                            if (!DATAG[0].Equals("PQC"))
                                dbShowData.InsertTAG(PTAG, txtProdNo.Text.ToUpper(), txtQCNo.Text, QtyT, txtOfTAG.Text, "PD", LineName, "PDTAG", PTAG2);
                        }
                    }
                }
                else if (FormISO.Equals("FM-PD-112")||FormISO.Equals("FM-PD-113") || FormISO.Equals("FM-PD-123") || FormISO.Equals("FM-PD-153") || FormISO.Equals("FM-PD-010")
                    || FormISO.Equals("FM-PD-164")
                    )
                {
                    TypeP = "PD";
                    decimal QtyT = 0;
                    decimal.TryParse(txtQtyofTAG.Text, out QtyT);
                    if (QtyT == 0)
                        QtyT = 1;
                    if (!txtOfTAG.Text.Equals(""))
                    {
                        string[] DATAG = PTAG.Split(',');
                        if (DATAG.Length == 8)
                        {
                            if (!DATAG[0].Equals("PQC"))
                                dbShowData.InsertTAG(PTAG, txtProdNo.Text.ToUpper(), txtQCNo.Text, QtyT, txtOfTAG.Text, "PD", LineName, "PDTAG", PTAG2);
                        }
                    }
                }
                else
                {

                }
            }
            catch { }
        }
        string UserAPP = "";
        private void LoadData()
        {
            try
            {
                //gobalNo = 0;
                //gobalNo2 = 0;
                radPageView1.Enabled = true;
                btnSave.Enabled = true;                
                btnAddNG.Enabled = true;
                txtScanTAG.Text = PTAG;
                if (!WOs.Equals(""))
                {
                    try
                    {
                        string[] Tag = PTAG.Split(',');
                        if (Tag.Length == 8)
                        {
                            txtQtyofTAG.Text = Tag[2].ToString();
                            txtOfTAG.Text = Tag[5].ToString();
                        }
                        else
                        {
                            txtOfTAG.Text = "1of1";
                        }
                    }
                    catch { }

                    using (DataClasses1DataContext db = new DataClasses1DataContext())
                    {
                        var listUser = db.tb_Users.ToList();
                        if(listUser.Count>0)
                        {
                            

                            
                        }

                        //tb_User us = db.tb_Users.Where(u => u.UserID.Equals(dbClss.UserID)).FirstOrDefault();
                        //if(us!=null)
                        //{
                        //    UserAPP = "";
                        //    cboCheckBy1.Text = UserAPP;
                        //    cboCheckBy2.Text = UserAPP;
                        //    cboCheckBy3.Text = UserAPP;
                        //    cboCreateBy1.Text = UserAPP;
                        //    cboCreateBy2.Text = UserAPP;

                        //}
                        
                        var woList = db.sp_46_QCSelectWO_01(WOs).FirstOrDefault();
                        if (woList != null)
                        {
                            
                            txtPartNo.Text = woList.CODE.ToString();
                            txtProdNo.Text = woList.PORDER.ToString().ToUpper();
                            txtLineNo.Text = woList.BUMO.ToString();
                            txtLotNo.Text = woList.LotNo.ToString();
                            txtQty.Text = Convert.ToDecimal(woList.OrderQty).ToString("###,###.##");
                            txtDayNight.Text = woList.DayNight.ToString();

                          //  txtOfTAG.Text = "";// woList.PrintTAG;

                            string Tx = db.get_QC_FromISOGet01(FormISO, 0);
                            txtStatus.Text = "Waiting";
                            txtStatus.ForeColor = Color.Red;
                            txtLotNo.Text = woList.LotNo;
                            txtLotSize.Text = woList.LotSize.ToString();            
                            groupBox1.Text = "Detail Production Order -> " + FormISO + " " + Tx;
                            txtChangeProduct.Text = Convert.ToDouble(woList.ChangeModel).ToString("###,###.##");
                            if(txtNo1PartNo.Text.Equals("") && txtPartNo.Text.Length>7)
                            {
                                txtNo1PartNo.Text = dbClss.Right(txtPartNo.Text, 7);
                            }
                            //Check Sttatus//                           
                            var ListHD = db.sp_46_QCSelectWO_05(txtProdNo.Text, FormISO).ToList();
                            if (ListHD.Count > 0)
                            {
                                var rs1 = ListHD.FirstOrDefault();
                                txtStatus.Text = rs1.Status;
                                txtStatus.ForeColor = Color.Red;
                                txtQCNo.Text = rs1.QCNo.ToString();
                                db.sp_46_QCSelectWO_12_UpdateOKQty(txtQCNo.Text);
                                //cboCheckBy1.Text = rs1.CheckBy1;
                                //cboCheckBy2.Text = rs1.CheckBy2;
                                //cboCheckBy3.Text = rs1.CheckBy3;
                                //cboCreateBy1.Text = rs1.IssueBy;
                                //cboCreateBy2.Text = rs1.IssueBy2;

                                if (txtStatus.Text.Equals("Completed"))
                                {
                                    btnSave.Enabled = false;
                                    btnAddNG.Enabled = false;
                                    txtStatus.ForeColor = Color.DarkGreen;
                                }
                                else if (rs1.SS == 2)
                                {
                                    btnSave.Enabled = false;
                                    btnAddNG.Enabled = false;
                                    txtStatus.ForeColor = Color.OrangeRed;
                                }
                                else if (txtStatus.Text.Equals("Checking"))
                                {
                                    txtStatus.ForeColor = Color.OrangeRed;
                                }
                            

                               // LoadNo(rs1.QCNo);
                                LoadNGPoint();
                            }

                           
                            var ListGroupPart = db.sp_46_QCMaster_Select(FormISO, txtPartNo.Text).ToList();
                            if (ListGroupPart.Count > 0)
                            {
                                cboCheckGroupPart.AutoSizeDropDownToBestFit = true;
                                cboCheckGroupPart.DisplayMember = "StepPart";
                                cboCheckGroupPart.ValueMember = "id";
                                cboCheckGroupPart.DataSource = ListGroupPart;
                                txtSeq.Text = "0";

                            }else
                            {
                                cboCheckGroupPart.Text = "";
                                txtSeq.Text = "0";
                            }
                            LoadChecker();




                        }
                        if(!txtPartNo.Text.Equals(""))
                        {
                            var liqg = db.sp_46_QCMaster_Select(FormISO, txtPartNo.Text).ToList();
                            gobalNo2 = liqg.Count-1;
                            if (gobalNo2 < 0)
                                gobalNo2 = 0;
                        }
                    }
                }
            }
            catch(Exception ex) { MessageBox.Show(ex.Message); }

            txtDataBox.Focus();
        }
        private void LoadNGPoint()
        {
            try
            {
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    /*
                    int rowS1 = 0;
                    string QCN = txtQCNo.Text;
                    if (QCN.Equals(""))
                        QCN = txtTempNo.Text;
                    radGridView2.DataSource = null;
                    radGridView2.DataSource = db.sp_46_QCMaster_SelectNG2(QCN, txtTempNo.Text,PTAG);
                    foreach(var rd in radGridView2.Rows)
                    {
                        rowS1 += 1;
                        rd.Cells["No"].Value = rowS1;
                    }
                    db.sp_46_QCNGUpdate(QCN, PTAG);
                    tb_QCTAG qg = db.tb_QCTAGs.Where(q => q.BarcodeTag.Equals(PTAG) && q.QCNo.Equals(QCN)).FirstOrDefault();
                    if(qg!=null)
                    {
                        txtNGInTAG.Text = Convert.ToInt32(qg.NGofTAG).ToString("###,###");
                    }
                    */
                    tb_QCTAG qt = db.tb_QCTAGs.Where(q => q.BarcodeTag.Equals(PTAG) && q.QCNo.Equals(txtQCNo.Text)).FirstOrDefault();
                    if (qt != null)
                    {
                        txtNGInTAG.Text= Convert.ToInt32(qt.NGofTAG).ToString("###,###");
                    }
                    txtNGPontQ.Text = "";
                    txtNGPontQ.Text = Convert.ToInt32(db.get_QCNGPointQ(txtQCNo.Text,PTAG)).ToString("###");
                }
            }
            catch { }
        }
        private void LoadNo(string QCNo)
        {
            try
            {
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    tb_QCNGPoint N1 = db.tb_QCNGPoints.Where(n => n.SeqNo == 1 && n.QCNo.Equals(QCNo) && n.PTAG.Equals(PTAG)).FirstOrDefault();
                    if(N1!=null)
                    {
                        //txtNo1NG.Text = N1.NGQty;
                        //txtNo1Remark.Text = N1.PointRemark.ToString();
                        //if(!N1.NGQty.Equals("0") && !N1.NGQty.Equals(""))
                        //{
                        //    ckNo1Check.Checked = true;
                        //}
                    }
                  


                    //QCTAG//
                    //tb_QCTAG RdT = db.tb_QCTAGs.Where(t => t.BarcodeTag.Equals(PTAG) && t.QCNo.Equals(QCNo)).FirstOrDefault();
                    //if(RdT!=null)
                    //{
                    //    //txtInspector.Text = RdT.CheckBy;
                    //    //dtDateInsp.Value = Convert.ToDateTime(RdT.CheckDate);
                    //    //txtNGALL.Text = Convert.ToString(RdT.NGQty);

                    //}

                    //QC SetupPoint//
                    tb_QCSetupPoint Sp1 = db.tb_QCSetupPoints.Where(s => s.RNo.Equals(1) && s.FormISO.Equals(FormISO) && s.WONo.Equals(txtProdNo.Text)).FirstOrDefault();
                    if(Sp1!=null)
                    {                        
                        txtNo1PartNo.Text = Sp1.Value1;
                    }
                    


                }
            }
            catch { }
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
            LoadData();
           

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
            SaveData();
        }
        private void SaveData()
        {
            try
            {
                if (MessageBox.Show("ต้องการบันทึกหรือไม่ ?", "บันทึก", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    using (DataClasses1DataContext db = new DataClasses1DataContext())
                    {

                        string QCP = "";


                        tb_QCHD qch = db.tb_QCHDs.Where(q => q.WONo.Equals(txtProdNo.Text) && q.FormISO.Equals(FormISO)).FirstOrDefault();
                        if (qch == null)
                        {
                            return;
                            txtQCNo.Text= dbClss.GetSeriesNo(6, 2);
                            tb_QCHD qcN = new tb_QCHD();
                            qcN.CheckBy1 = "";
                            qcN.CheckBy2 = "";
                            qcN.CheckBy3 = "";
                            qcN.IssueBy = "";
                            qcN.IssueBy2 = "";
                            qcN.ApproveBy = "";
                            qcN.ApproveBy2 = "";
                            qcN.ChangeModel = 0;
                            qcN.QCNo = txtQCNo.Text;
                            qcN.WONo = txtProdNo.Text.ToUpper();
                            qcN.PartNo = txtPartNo.Text;
                            qcN.OrderQty = Convert.ToDecimal(txtQty.Text);
                            qcN.OKQty = 0;
                            qcN.NGQty = 0;
                            qcN.LotNo = txtLotNo.Text;
                            qcN.LineName = LineName;
                            qcN.CreateBy = dbClss.UserID;
                            qcN.CreateDate = DateTime.Now;
                            qcN.SS = 1;
                            qcN.Status = "Checking";
                            qcN.PassValue = "";
                            qcN.PDQty = 0;
                            qcN.SendApprove = false;
                            if (chkApprove.Checked)
                            {
                                qcN.Status = "Waiting Approve";
                                qcN.SS = 2;
                                qcN.SendApprove = true;
                            }
                            qcN.FormISO = FormISO;
                            qcN.DocRef1 = "";
                            qcN.DocRef2 = "";
                            qcN.ApproveBy = "";
                            qcN.ApproveBy2 = "";
                            //qcN.CheckBy1 = cboCheckBy1.Text;
                            //qcN.CheckBy2 = cboCheckBy2.Text;
                            //qcN.CheckBy3 = cboCheckBy3.Text;
                         
                           // qcN.IssueBy = cboCreateBy1.Text;
                            qcN.IssueDate = DateTime.Now;
                          //  qcN.IssueBy2 = cboCreateBy2.Text;
                            qcN.IssueDate2 = DateTime.Now;
                            qcN.ChangeModel = Convert.ToDecimal(txtChangeProduct.Text);
                            qcN.DayNight = txtDayNight.Text;

                           // if (!cboCheckBy1.Text.Equals(""))
                                qcN.CheckDate1 = DateTime.Now;
                         //   if (!cboCheckBy2.Text.Equals(""))
                                qcN.CheckDate2 = DateTime.Now;
                          //  if (!cboCheckBy3.Text.Equals(""))
                                qcN.CheckDate3 = DateTime.Now;                                                      
                            
                            qcN.QCPoint = QCP;
                            db.tb_QCHDs.InsertOnSubmit(qcN);
                            db.SubmitChanges();
                            
                        }
                        else
                        {
                            
                            if (!Convert.ToBoolean(qch.SendApprove))
                            {
                                if (chkApprove.Checked)
                                {
                                    qch.SendApprove = true;
                                    qch.SS = 2;
                                    qch.Status = "Waiting Approve";
                                }


                            }
                            db.SubmitChanges();
                            
                        }


                        //Insert SetUpNGPoint//
                        //  InsertSetupNGPoint(1, "Part No.",txtNo1PartNo.Text);                       
                        // InsertSetupNGPoint(41, "SERVICE",txtNo4Service.Text);
                        //  InsertSetupNGPoint(42, "EMERGENCY", txtNo4Emergency.Text);
                        //   InsertSetupNGPoint(43, "Lot No.", txtLotNo.Text);                        


                        //Insert NGPoint//
                        tb_QCHD Upd = db.tb_QCHDs.Where(q => q.WONo.Equals(txtProdNo.Text) && q.FormISO.Equals(FormISO)).FirstOrDefault();
                        if (Upd != null)
                        {



                            //////////Insert NG////////////
                            //db.sp_46_QCHD_copy(txtQCNo.Text, txtTempNo.Text);
                            
                          //  int OK = 0;
                          //  int NG = 0;
                           // int NG1, NG2, NG3, NG4, NG5, NG6 = 0;

                            //SumQty Inspection / NG / OK 
                            /*
                            decimal SumALL = Convert.ToDecimal(db.get_QCSumQty(Upd.QCNo, PTAG, 5));
                            decimal SumOK = 0;// Convert.ToDecimal(db.get_QCSumQty(Upd.QCNo, 1));
                            decimal SumNG = Convert.ToDecimal(db.get_QCSumQty(Upd.QCNo, PTAG, 4));
                            SumOK = SumALL - SumNG;


                            tb_QCTAG qctag = db.tb_QCTAGs.Where(t => t.BarcodeTag.Equals(PTAG) && t.QCNo.Equals(Upd.QCNo)).FirstOrDefault();
                            if (qctag == null)
                            {
                                tb_QCTAG qct = new tb_QCTAG();
                                qct.QCNo = Upd.QCNo;
                                qct.BarcodeTag = PTAG;
                                qct.SS = 1;
                                qct.QtyofTag = Convert.ToDecimal(txtQty.Text);
                                qct.OKQty = (OK - NG);
                                qct.NGQty = NG;
                                qct.ofTAG = txtOfTAG.Text;
                                qct.Dept = "PD";
                                qct.CheckDate = DateTime.Now;
                                qct.CheckBy = cboCheckBy1.Text;
                                qct.DType = txtLineNo.Text;
                                db.tb_QCTAGs.InsertOnSubmit(qct);
                                db.SubmitChanges();
                            }
                            else
                            {

                                qctag.QtyofTag = Convert.ToDecimal(txtQty.Text);
                                qctag.OKQty = (OK - NG);
                                qctag.NGQty = NG;
                                qctag.ofTAG = txtOfTAG.Text;
                                db.SubmitChanges();
                            }



                            //Update Upd HD//
                            Upd.OrderQty = SumALL;
                            Upd.NGQty = SumNG;
                            Upd.OKQty = SumOK;
                            */
                            if (Upd.SendApprove.Equals(false) && chkApprove.Checked)
                            {
                                Upd.SendApprove = chkApprove.Checked;
                                if (Upd.SS == 1)
                                    Upd.SS = 2;
                            }
                            db.SubmitChanges();

                            if (chkApprove.Checked)
                            {
                                if (CheckApprove())
                                {
                                    tb_QCApprove qa = db.tb_QCApproves.Where(w => w.WONo.Equals(Upd.WONo) && w.FormISO.Equals(Upd.FormISO)).FirstOrDefault();
                                    if (qa == null)
                                    {
                                        tb_QCApprove ap = new tb_QCApprove();
                                        ap.FormISO = Upd.FormISO;
                                        ap.WONo = Upd.WONo;
                                        ap.PartName = "";
                                        ap.PartNo = Upd.PartNo;
                                        ap.Seq = 1;
                                        ap.Remark = "";
                                        ap.OKQty = 0;
                                        ap.NGQty = 0;
                                        ap.InsQty = 0;
                                        ap.ApproveBy = "";
                                        ap.ApproveDate = null;
                                        ap.SS = 1;
                                        ap.QCNo = txtQCNo.Text;
                                        db.tb_QCApproves.InsertOnSubmit(ap);
                                        db.SubmitChanges();



                                    }
                                    else
                                    {

                                    }
                                }
                            }


                            //CallStore for Update Qty,OK,NG////
                            db.sp_46_QCHD_Update_HD(txtQCNo.Text.ToUpper());


                        }


                    }

                    MessageBox.Show("บันทึกสำเร็จ");
                    LoadData();
                }

            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }
        private bool CheckApprove()
        {
            bool ck = true;
            using (DataClasses1DataContext db = new DataClasses1DataContext())
            {
                if (!rdSkipCheck.Checked)
                {
                    if (FormISO.Equals("FM-QA-055") || FormISO.Equals("FM-QA-055_02_1"))
                    {
                        var PDRC = db.tb_ProductionReceives.Where(p => p.OrderNo.Equals(txtProdNo.Text)).ToList();
                        var QCTag = db.tb_QCTAGs.Where(p => p.QCNo.Equals(txtQCNo.Text)).ToList();
                        if (PDRC.Count > QCTag.Count)
                        {
                            ck = false;
                            MessageBox.Show("จำนวน TAG ที่ Scan เช็คยังไม่ครบตาม จาก PD");
                            tb_QCHD qch = db.tb_QCHDs.Where(q => q.WONo.Equals(txtProdNo.Text) && q.FormISO.Equals(FormISO)).FirstOrDefault();
                            if (qch != null)
                            {
                                if(!rdSkipCheck.Checked)
                                {
                                    qch.Status = "Checking";
                                    db.SubmitChanges();
                                }
                                
                            }
                        }
                    }
                }else
                {
                   
                    tb_QCHD qch = db.tb_QCHDs.Where(q => q.WONo.Equals(txtProdNo.Text) && q.FormISO.Equals(FormISO)).FirstOrDefault();
                    if (qch != null)
                    {
                        qch.Status = "Waiting Approve";
                        db.SubmitChanges();
                    }
                }
            }
            return ck;
        }

        private void InsertNGPoint(int No1,string QCNo,string PointName,string PointRemark,string PointRemark2,string ofTAG,string TopCaseText,string NGQty)
        {
            try
            {
                using (DataClasses1DataContext db2 = new DataClasses1DataContext())
                {
                    decimal NGQ = 0;
                    decimal.TryParse(NGQty, out NGQ);
                    tb_QCNGPoint cki = db2.tb_QCNGPoints.Where(qq => qq.SeqNo.Equals(No1) && qq.PTAG.Equals(PTAG) && qq.QCNo.Equals(QCNo)).FirstOrDefault();
                    if (cki != null)
                    {
                        cki.PointName = PointName;
                        cki.PointRemark = PointRemark;
                        cki.NGQty = NGQ.ToString();
                        db2.SubmitChanges();
                    }
                    else
                    {
                        tb_QCNGPoint qcn = new tb_QCNGPoint();
                        qcn.QCNo = QCNo;
                        qcn.Status = "Waiting";
                        qcn.SeqNo = No1;
                        qcn.PTAG = PTAG;
                        qcn.TopCaseText = TopCaseText;
                        qcn.PointName = PointName;
                        qcn.PointRemark = PointRemark;
                        qcn.PointRemark2 = PointRemark2;
                        qcn.ofTag = ofTAG;
                        qcn.NCRNo = "";
                        qcn.NCRSS = 0;
                        qcn.NGQty = Convert.ToString(NGQ);
                        qcn.WONo = txtProdNo.Text.ToUpper();
                        db2.tb_QCNGPoints.InsertOnSubmit(qcn);
                        db2.SubmitChanges();

                    }
                }
            }
            catch { }
        }
        private void InsertSetupNGPoint(int No1,string Desc,string Value1)
        {
            try
            {
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    tb_QCSetupPoint sq = db.tb_QCSetupPoints.Where(s => s.FormISO.Equals(FormISO) && s.WONo.Equals(txtProdNo.Text) && s.RNo.Equals(No1)).FirstOrDefault();
                    if (sq != null)
                    {
                        sq.Value1 = Value1;
                        db.SubmitChanges();
                    }
                    else
                    {
                        tb_QCSetupPoint sn = new tb_QCSetupPoint();
                        sn.WONo = txtProdNo.Text.ToUpper();
                        sn.FormISO = FormISO;
                        sn.RNo = No1;
                        sn.Seq = db.get_MaxSetupNGPoint(txtProdNo.Text.ToUpper(), FormISO, 0)+1;
                        sn.Value1 = Value1;
                        sn.Description = Desc;
                        sn.Rid = 0;
                        db.tb_QCSetupPoints.InsertOnSubmit(sn);
                        db.SubmitChanges();
                    }
                }
            }
            catch { }
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
        private void getWO()
        {
            try
            {

              
                    using (DataClasses1DataContext db = new DataClasses1DataContext())
                    {
                       
                        string WO = "";                

                        var woList = db.sp_46_QCSelectWO_01(WO.ToUpper()).FirstOrDefault();
                        

                        txtPartNo.Text = woList.CODE.ToString();
                        txtProdNo.Text = woList.PORDER.ToString();
                        txtLineNo.Text = woList.BUMO.ToString();
                        txtQty.Text = Convert.ToDecimal(woList.OrderQty).ToString("###,###.##");
                        txtLotNo.Text = woList.LotNo.ToString();
                        var FormList = db.sp_46_QCSelectWO_02(txtProdNo.Text.ToUpper(), LineName, txtPartNo.Text,"PD").ToList();
                      //  radGridView2.DataSource = FormList;
                       
                        ////Load Datagridview///
                    }
                
            }
            catch { }
        }

        int PrintC = 0;
        private void radButtonElement1_Click(object sender, EventArgs e)
        {
            //PrintData(txtProdNo.Text,txtPartNo.Text);
            //PrintC = 0;
            this.Cursor = Cursors.WaitCursor;
            try
            {
                if (PrintC == 0)
                {
                    PrintC = 1;
                    dbShowData.CallReportQC(txtProdNo.Text.ToUpper(),
                        txtPartNo.Text.ToUpper(),
                        txtQCNo.Text,
                       FormISO);


                    //if (FormISO.Equals("FM-PD-026_1"))
                    //{
                    //    dbShowData.PrintData(txtProdNo.Text, txtPartNo.Text, txtQCNo.Text);
                    //}
                    //else if (FormISO.Equals("FM-PD-033_1"))
                    //{
                    //    dbShowData.PrintData033(txtProdNo.Text, txtPartNo.Text, txtQCNo.Text);
                    //}
                    //else if (FormISO.Equals("FM-PD-035_1"))
                    //{
                    //    dbShowData.PrintData035(txtProdNo.Text.ToUpper(), txtPartNo.Text, txtQCNo.Text);
                    //}
                    //else if (FormISO.Equals("FM-QA-055"))
                    //{
                    //    dbShowData.PrintData55CT(txtProdNo.Text.ToUpper(), txtPartNo.Text, txtQCNo.Text);
                    //}
                    //else if (FormISO.Equals("FM-QA-055_02_1"))
                    //{
                    //    dbShowData.PrintData5501(txtProdNo.Text.ToUpper(), txtPartNo.Text, txtQCNo.Text);
                    //}
                    //else if (FormISO.Equals("FM-QA-056_02_1"))
                    //{
                    //    dbShowData.PrintData5601(txtProdNo.Text.ToUpper(), txtPartNo.Text, txtQCNo.Text);
                    //}else if(FormISO.Equals("FM-QA-091")
                    //    ||FormISO.Equals("FM-QA-092")
                    //    || FormISO.Equals("FM-QA-143")
                    //    || FormISO.Equals("FM-QA-144")
                    //    || FormISO.Equals("FM-QA-161")
                    //    )
                    //{
                    //    dbShowData.PrintData56CT(txtProdNo.Text.ToUpper(), txtPartNo.Text, txtQCNo.Text);
                    //}
                    //else if (FormISO.Equals("FM-PD-001"))
                    //{
                    //    dbShowData.PrintFMPD001(txtProdNo.Text.ToUpper(), txtPartNo.Text, txtQCNo.Text,FormISO);
                    //}
                    //else if (FormISO.Equals("FM-PD-002"))
                    //{
                    //    dbShowData.PrintFMPD002(txtProdNo.Text.ToUpper(), txtPartNo.Text, txtQCNo.Text, FormISO);
                    //}
                    //else if (FormISO.Equals("FM-PD-003"))
                    //{
                    //    dbShowData.PrintFMPD003(txtProdNo.Text.ToUpper(), txtPartNo.Text, txtQCNo.Text, FormISO);
                    //}
                    //else if (FormISO.Equals("FM-PD-109"))
                    //{
                    //    dbShowData.PrintFMPD109(txtProdNo.Text.ToUpper(), txtPartNo.Text, txtQCNo.Text, FormISO);
                    //}
                    //else if (FormISO.Equals("FM-PD-095"))
                    //{
                    //    dbShowData.PrintFMPD095(txtProdNo.Text.ToUpper(), txtPartNo.Text, txtQCNo.Text, FormISO);
                    //}
                    //else if (FormISO.Equals("FM-PD-096"))
                    //{
                    //    dbShowData.PrintFMPD096(txtProdNo.Text.ToUpper(), txtPartNo.Text, txtQCNo.Text, FormISO);
                    //}
                    //else if (FormISO.Equals("FM-PD-003_S")||FormISO.Equals("FM-PD-156"))
                    //{
                    //    dbShowData.PrintFMPD003_S(txtProdNo.Text.ToUpper(), txtPartNo.Text, txtQCNo.Text, FormISO);
                    //}
                    //else if (FormISO.Equals("FM-PD-112"))
                    //{
                    //    dbShowData.PrintFMPD112(txtProdNo.Text.ToUpper(), txtPartNo.Text, txtQCNo.Text, FormISO);
                    //}
                    //else if (FormISO.Equals("FM-PD-110"))
                    //{
                    //    dbShowData.PrintPD110(txtProdNo.Text.ToUpper(), txtPartNo.Text, txtQCNo.Text, FormISO);
                    //}
                    //else if (FormISO.Equals("FM-PD-157"))
                    //{
                    //    dbShowData.PrintPD157(txtProdNo.Text.ToUpper(), txtPartNo.Text, txtQCNo.Text, FormISO);
                    //}
                    //else if (FormISO.Equals("FM-PD-077"))
                    //{
                    //    dbShowData.PrintPD077(txtProdNo.Text.ToUpper(), txtPartNo.Text, txtQCNo.Text, FormISO);
                    //}
                    //else if (FormISO.Equals("FM-PD-113"))
                    //{
                    //    dbShowData.PrintPD113(txtProdNo.Text.ToUpper(), txtPartNo.Text, txtQCNo.Text, FormISO);
                    //}
                    //else if (FormISO.Equals("FM-PD-123"))
                    //{
                    //    dbShowData.PrintPD123(txtProdNo.Text.ToUpper(), txtPartNo.Text, txtQCNo.Text, FormISO);
                    //}

                }
            }
            catch { PrintC = 0; }
            PrintC = 0;
                this.Cursor = Cursors.Default;
            System.Threading.Thread.Sleep(1000);

        }

        public void PrintData(string WO,string PartNo)
        {
            try
            {
                /*
                string FormISOx = "FM-PD-026_00_1.rpt";
                Report.Reportx1.WReport = "QCReport01";
                Report.Reportx1.Value = new string[2];
                Report.Reportx1.Value[0] = FormISOx;
                Report.Reportx1.Value[1] = txtProdNo.Text;
                Report.Reportx1 op = new Report.Reportx1(FormISOx);
                op.Show();
                */
                this.Cursor = Cursors.WaitCursor;
                string DATA = AppDomain.CurrentDomain.BaseDirectory;
                string tempPath = System.IO.Path.GetTempPath();
                string FileName = "FM-PD-026_1.xlsx";
                string tempfile = tempPath + FileName;
                DATA = DATA + @"QC\" + FileName;

                if (File.Exists(tempfile))
                {
                    try
                    {
                        File.Delete(tempfile);
                    }
                    catch { }
                }

                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook excelBook = excelApp.Workbooks.Open(
                  DATA, 0, true, 5,
                  "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false,
                  0, true);
                Excel.Sheets sheets = excelBook.Worksheets;
                Excel.Worksheet worksheet = (Excel.Worksheet)sheets.get_Item(1);

                // progressBar1.Maximum = 51;
                // progressBar1.Minimum = 1;
                int row1 = 17;
                int row2 = 17;
                int Seq = 0;
                int seq2 = 21;
                int CountRow = 0;
                string PV = "P";
                string QHNo = "";
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    string Value1 = "";
                    string Value2 = "";
                    string LotNo = "";
                    ///////////////SETValue/////////////////
                    var DValue = db.sp_46_QCSelectWO_01(WO).FirstOrDefault();
                    if (DValue != null)
                    {

                        Excel.Range CPart = worksheet.get_Range("P3");
                        CPart.Value2 = DValue.CODE;

                        Excel.Range CStamp = worksheet.get_Range("X5");
                        CStamp.Value2 = dbClss.Right(PartNo, 7);

                        

                        Excel.Range CName = worksheet.get_Range("I5");
                        CName.Value2 = DValue.NAME;

                        Excel.Range CDate = worksheet.get_Range("D5");
                        CDate.Value2 = DValue.DeliveryDate;

                        Excel.Range CLot = worksheet.get_Range("D7");
                        CLot.Value2 = DValue.LotNo;

                        Excel.Range CQty = worksheet.get_Range("D9");
                        CQty.Value2 = DValue.OrderQty.ToString();

                        
                        try
                        {
                            tb_QCHD qh = db.tb_QCHDs.Where(w => w.QCNo.Equals(txtQCNo.Text)).FirstOrDefault();
                            if (qh != null)
                            {
                                Excel.Range Ap = worksheet.get_Range("AE12");
                                Ap.Value2 = Convert.ToString(qh.ApproveBy);

                                Excel.Range CheckBy1 = worksheet.get_Range("E18");
                                CheckBy1.Value2 = qh.CheckBy1;
                                Excel.Range CheckBy2 = worksheet.get_Range("E28");
                                CheckBy2.Value2 = qh.CheckBy2;
                                Excel.Range CheckBy3 = worksheet.get_Range("E34");
                                CheckBy3.Value2 = qh.CheckBy3;
                                Excel.Range IssueBy = worksheet.get_Range("AE3");
                                IssueBy.Value2 = qh.IssueBy;

                                QHNo = qh.QCNo;
                            }
                            var gTime = db.sp_46_QCGetValue2601_Time(WO).ToList();
                            if (gTime.Count > 0)
                            {
                                var g = gTime.FirstOrDefault();
                                Excel.Range AB = worksheet.get_Range("AB9");
                                AB.Value2 = Convert.ToDecimal(DValue.ChangeModel).ToString("####") + " นาที";

                                if (!g.StartTime.Equals(""))
                                {
                                    Excel.Range StartT = worksheet.get_Range("N7");
                                    StartT.Value2 = Convert.ToDateTime(g.StartTime).ToString("HH:mm");

                                    Excel.Range EndT = worksheet.get_Range("AA7");
                                    EndT.Value2 = Convert.ToDateTime(g.EndTime).ToString("HH:mm");

                                    int ChanP = 0;
                                    int.TryParse(Convert.ToInt32(DValue.ChangeModel).ToString(), out ChanP);
                                    if (ChanP > 0)
                                    {
                                        DateTime Chtime = Convert.ToDateTime(g.StartTime).AddMinutes(ChanP * -1);
                                        Excel.Range O9 = worksheet.get_Range("O9");
                                        O9.Value2 = "'" + Convert.ToDateTime(Chtime).ToString("HH:mm") + "-" + Convert.ToDateTime(g.StartTime).ToString("HH:mm");

                                    }

                                }
                            }
                        }
                        catch { }




                    }

                    ////////////////////////////////////////


                    var listPart = db.tb_QCGroupParts.Where(q => q.FormISO.Equals(FormISO) && q.PartNo.Equals(PartNo)).OrderBy(o => o.Seq).ToList();
                    foreach (var rd in listPart)
                    {

                        if (CountRow == 0)
                        {
                            if (rd.Seq.Equals(48))
                            {
                                Excel.Range CRemark = worksheet.get_Range("A13");
                                CRemark.Value2 = "Remark  " + rd.SetData;
                                CountRow += 1;
                            }
                        }

                        if (rd.Seq < 22)
                        {
                            row1 += 1;
                            Seq += 1;
                            if (row1 <= 38)
                            {

                                Excel.Range Col0 = worksheet.get_Range("G" + row1.ToString(), "G" + row1.ToString());
                                Excel.Range Col1 = worksheet.get_Range("L" + row1.ToString(), "L" + row1.ToString());
                                if (Seq.Equals(rd.Seq))
                                {
                                    Col0.Value2 = rd.TopPic;
                                    Col1.Value2 = rd.SetData;
                                    if (!rd.SetData.Equals(""))
                                    {
                                        try
                                        {
                                            var gValue = db.sp_46_QCGetValue2601(WO, rd.SetData).FirstOrDefault();

                                            LotNo = "";
                                            LotNo = Convert.ToString(gValue.Lot);
                                            if (gValue.CountA > 0)
                                            {
                                                if (txtDayNight.Text.Equals("D"))
                                                {
                                                    Excel.Range Check1 = worksheet.get_Range("Q" + row1.ToString(), "Q" + row1.ToString());
                                                    Check1.Value2 = "P";
                                                }
                                                else
                                                {
                                                    Excel.Range Check2 = worksheet.get_Range("R" + row1.ToString(), "R" + row1.ToString());
                                                    Check2.Value2 = "P";
                                                }

                                                if (!LotNo.Equals(""))
                                                {
                                                    Excel.Range Check3 = worksheet.get_Range("S" + row1.ToString(), "S" + row1.ToString());
                                                    Check3.Value2 = LotNo;
                                                }
                                            }
                                        }
                                        catch (Exception ex) { MessageBox.Show(ex.Message); }
                                    }

                                }
                                if (row1 == 18)
                                    row1 += 1;
                            }
                        }
                        else
                        {
                            row2 += 1;
                            seq2 += 1;
                            PV = "P";
                            if (row2 == 25 || row2 == 43)
                                row2 += 1;
                            if (seq2.Equals(rd.Seq) && rd.Seq != 48)
                            {
                                if (row2 != 31 || row2 != 42)
                                {
                                    Excel.Range Col2 = worksheet.get_Range("AA" + row2.ToString(), "AA" + row2.ToString());
                                    Col2.Value2 = rd.TopPic;
                                }
                                if (row2 != 24 || row2 != 42)
                                {
                                    Excel.Range Col3 = worksheet.get_Range("AE" + row2.ToString(), "AE" + row2.ToString());
                                    Col3.Value2 = rd.SetData;

                                }

                                if (row2 != 42 && row2 != 43)
                                {
                                    tb_QCNGPoint ngq = db.tb_QCNGPoints.Where(w => w.QCNo.Equals(QHNo) && w.SeqNo.Equals(rd.Seq)).FirstOrDefault();
                                    if (ngq != null)
                                    {
                                        PV = "O";
                                    }

                                    if (txtDayNight.Text.Equals("D"))
                                    {
                                        Excel.Range Check2 = worksheet.get_Range("AF" + row2.ToString(), "AF" + row2.ToString());
                                        Check2.Value2 = PV;
                                    }
                                    else
                                    {
                                        Excel.Range Check2 = worksheet.get_Range("AG" + row2.ToString(), "AG" + row2.ToString());
                                        Check2.Value2 = PV;
                                    }

                                    if (row2 == 35)
                                    {
                                        Excel.Range Check2 = worksheet.get_Range("AG" + row2.ToString(), "AG" + row2.ToString());
                                        Check2.Value2 = rd.SetData;
                                    }
                                }



                            }
                        }



                    }

                    /*
                    for (int j = 0; j <= 50; j++)
                    {
                        row1 += 1;
                        Excel.Range Col0 = worksheet.get_Range("B" + row1.ToString(), "B" + row1.ToString());
                        // Excel.Range Col1 = worksheet.get_Range("E" + row1.ToString(), "E" + row1.ToString());
                        Excel.Range Col2 = worksheet.get_Range("F" + row1.ToString(), "F" + row1.ToString());
                        Excel.Range Col3 = worksheet.get_Range("C" + row1.ToString(), "C" + row1.ToString());
                        string Value1 = Convert.ToString(Col0.Value2);
                        if (Value1 == null)
                        {
                            Value1 = "";
                        }
                        if (!Convert.ToString(Value1).Equals(""))
                        {
                            Seq = 0;
                            int.TryParse(Value1, out Seq);
                            Col2.Value = db.QC_GetTemplate(FormISO, txtPartNo.Text, Seq);
                            Col3.Value = txtPartNo.Text.ToUpper();

                        }

                    }
                    */
                }

                excelBook.SaveAs(tempfile);
                excelBook.Close(false);
                excelApp.Quit();

                releaseObject(worksheet);
                releaseObject(excelBook);
                releaseObject(excelApp);
                Marshal.FinalReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                GC.GetTotalMemory(false);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.GetTotalMemory(true);
                System.Diagnostics.Process.Start(tempfile);

            }
            catch { }
            this.Cursor = Cursors.Default;
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
        private void radButton2_Click_1(object sender, EventArgs e)
        {
            NextPage(radPageViewPage7);
        }

        private void NextPage(RadPageViewPage Pv2)
        {
            Pv2.Enabled = true;
            radPageView1.SelectedPage = Pv2;
            Snew += 1;
            // MessageBox.Show(radPageView1.SelectedPage.Name);
            if (radPageView1.SelectedPage.Name.Equals("radPageViewPage1"))
            {
                txtInspector.Focus();
            }
            else
            {
                txtDataBox.Focus();
            }
        }

        private void btnBackNo7_Click(object sender, EventArgs e)
        {
            NextPage(radPageViewPage1);
        }

        private void btnNextNo6_Click(object sender, EventArgs e)
        {
           
        }

        private void btnBackNo6_Click(object sender, EventArgs e)
        {
           
        }

        private void btnNextNo5_Click(object sender, EventArgs e)
        {
            
        }

        private void btnBackNo5_Click(object sender, EventArgs e)
        {
           
        }

        private void btnNextNo4_Click(object sender, EventArgs e)
        {
           
        }

        private void btnBackNo4_Click(object sender, EventArgs e)
        {
           
        }

        private void txtNextNo3_Click(object sender, EventArgs e)
        {
           
        }

        private void txtBackNo3_Click(object sender, EventArgs e)
        {
           
        }

        private void btnNextNo2_Click(object sender, EventArgs e)
        {
           
        }

        private void btnBackNo2_Click(object sender, EventArgs e)
        {
            NextPage(radPageViewPage1);
        }

        private void radPageViewPage7_Paint(object sender, PaintEventArgs e)
        {
            try
            {
               
                SumNGALL();
            }
            catch { }
        }
        private void SumNGALL()
        {
            try
            {
                if (!txtQty.Text.Equals(""))
                {

                    int OK = 0;
                    int NG = 0;
                    int NG1, NG2, NG3, NG4, NG5, NG6 = 0;
                    //int.TryParse(txtNo1NG.Text, out NG1);
                    //int.TryParse(txtNo2NG.Text, out NG2);
                    //int.TryParse(txtNo3NG.Text, out NG3);
                    //int.TryParse(txtNo4NG.Text, out NG4);
                    //int.TryParse(txtNo5NG.Text, out NG5);
                    //int.TryParse(txtNo6NG.Text, out NG6);
                    //int.TryParse(txtNGALL.Text, out NG);
                    int.TryParse(txtQty.Text, out OK);
                 //   int ALLNG = (NG1 + NG2 + NG3 + NG4 + NG5 + NG6);

                    //if (ALLNG > OK)
                    //{
                    //    ALLNG = OK;
                    //}

                    //if (ckNo1Check.Checked || ckNo2Check.Checked || ckNo3Check.Checked || ckNo4Check.Checked ||  ckNo5Check.Checked || ckNo6Check.Checked || ckNo7Check.Checked)
                    //{
                    //    ALLNG = 1;
                    //}else
                    //{
                    //    ALLNG = 0;
                    //}
                  //if (txtNGALL.Text.Equals("") || txtNGALL.Text.Equals("0"))
                  //      txtNGALL.Text = (ALLNG).ToString();
                }
               
            }
            catch { }
        }

        private void btnBackNo8_Click(object sender, EventArgs e)
        {
           // NextPage(radPageViewPage6);
        }

        private void btnNextNo8_Click(object sender, EventArgs e)
        {
            NextPage(radPageViewPage7);
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                DeleteNG(0);
            }
            catch { }
        }
        private void DeleteNG(int id)
        {
            try
            {
                if (rowNG >= 0 && !txtStatus.Text.Equals("Completed"))
                {
                    if (MessageBox.Show("ต้องการลบหรือไม่", "ต้องการลบ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        int idng = 0;
                        int.TryParse(radGridView2.Rows[rowNG].Cells["id"].Value.ToString(), out idng);
                        if (idng > 0)
                        {
                            using (DataClasses1DataContext db = new DataClasses1DataContext())
                            {
                                string QCNo = txtQCNo.Text;
                                if (QCNo.Equals(""))
                                    QCNo = txtTempNo.Text;
                                db.sp_46_QCSelectWO_10_DeleteNGPoint(QCNo, idng);
                                MessageBox.Show("ลบเรียบร้อย");
                                LoadNGPoint();
                                //tb_QCNGPoint qng = db.tb_QCNGPoints.Where(q => q.id.Equals(idng) && !q.Status.Equals("Completed")).FirstOrDefault();
                                //if (qng != null)
                                //{
                                //    db.tb_QCNGPoints.DeleteOnSubmit(qng);
                                //    db.SubmitChanges();
                                //    MessageBox.Show("ลบเรียบร้อย");
                                //    LoadNGPoint();

                                //}
                            }
                        }
                    }
                }
            }
            catch { }
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {

        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {

        }

        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {

        }
        string globalPath = "";
        string globalFile = "";
        private void radButton1_Click_1(object sender, EventArgs e)
        {
            ShowImg();
            return;
            //image1
            try
            {
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    tb_QCImageCheck im = db.tb_QCImageChecks.Where(u => u.PartNo.Equals(txtPartNo.Text)).FirstOrDefault();
                    tb_Path ph = db.tb_Paths.Where(p => p.PathCode.Equals("QC1")).FirstOrDefault();
                    string Path = "";
                    if (ph != null)
                    {
                        Path = ph.PathFile;
                    }
                    if (im != null)
                    {
                        if (!im.Image1.Equals(""))
                        {
                            pictureBox1.Image = Image.FromFile(Path + im.Image1);
                            globalPath = Path + im.Image1;
                        }
                    }
                }

            }
            catch { }
        }

        private void radButton2_Click_2(object sender, EventArgs e)
        {
            //image2
            try
            {
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    tb_QCImageCheck im = db.tb_QCImageChecks.Where(u => u.PartNo.Equals(txtPartNo.Text)).FirstOrDefault();
                    tb_Path ph = db.tb_Paths.Where(p => p.PathCode.Equals("QC2")).FirstOrDefault();
                    string Path = "";
                    if (ph != null)
                    {
                        Path = ph.PathFile;
                    }
                    if (im != null)
                    {
                        if (!im.Image2.Equals(""))
                        {
                            pictureBox1.Image = Image.FromFile(Path + im.Image2);
                            globalPath = Path + im.Image2;
                        }
                    }
                }

            }
            catch { }
        }

        private void radButton3_Click_1(object sender, EventArgs e)
        {
            //image3
            try
            {
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    tb_QCImageCheck im = db.tb_QCImageChecks.Where(u => u.PartNo.Equals(txtPartNo.Text)).FirstOrDefault();
                    tb_Path ph = db.tb_Paths.Where(p => p.PathCode.Equals("QC3")).FirstOrDefault();
                    string Path = "";
                    if (ph != null)
                    {
                        Path = ph.PathFile;
                    }
                    if (im != null)
                    {
                        if (!im.Image3.Equals(""))
                        {
                            pictureBox1.Image = Image.FromFile(Path + im.Image3);
                            globalPath = Path + im.Image3;
                        }
                    }
                }

            }
            catch { }
        }

        private void radButton4_Click(object sender, EventArgs e)
        {
            //image4
            try
            {
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    tb_QCImageCheck im = db.tb_QCImageChecks.Where(u => u.PartNo.Equals(txtPartNo.Text)).FirstOrDefault();
                    tb_Path ph = db.tb_Paths.Where(p => p.PathCode.Equals("QC4")).FirstOrDefault();
                    string Path = "";
                    if (ph != null)
                    {
                        Path = ph.PathFile;
                    }
                    if (im != null)
                    {
                        if (!im.Image4.Equals(""))
                        {
                            pictureBox1.Image = Image.FromFile(Path + im.Image4);
                            globalPath = Path + im.Image4;
                        }
                    }
                }

            }
            catch { }
        }

        private void cboCheckGroupPart_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {

                txtNGID.Text = cboCheckGroupPart.SelectedValue.ToString();
                LoadCheckGroupPart(txtNGID.Text);
               
            }
            catch { }
        }
        private void LoadCheckGroupPart(string idg)
        {
            try
            {
                if (idg != "0" || idg !="")
                {

                    using (DataClasses1DataContext db = new DataClasses1DataContext())
                    {
                        int idQg = 0;
                        int.TryParse(idg, out idQg);
                        tb_QCGroupPart qg = db.tb_QCGroupParts.Where(q => q.id.Equals(idQg)).FirstOrDefault();
                        if (qg != null)
                        {
                            txtToppic.Text = qg.TopPic;
                            txtRank.Text = Convert.ToString(qg.Rank);
                            txtSetData.Text = qg.SetData;
                            if (FormISO.Equals("FM-PD-035_1"))
                            {
                                if (qg.Seq.Equals(5))
                                {
                                    txtSetData.Text = txtSetData.Text + Environment.NewLine + LotMark;
                                }
                            }

                            lblSeq.Text = "ลำดับ " + Convert.ToString(qg.Seq);
                            /////new 31/07/2020
                            txtSeq.Text = Convert.ToString(qg.Seq);
                            txtNGQty.Text = "";
                            txtValue.Text = "";
                            tb_QCNGPoint ngq = db.tb_QCNGPoints.Where(q => q.QCNo.Equals(txtQCNo.Text) && q.SeqNo.Equals(qg.Seq) && q.PTAG.Equals(txtScanTAG.Text)).FirstOrDefault();
                            if (ngq != null)
                            {
                                txtValue.Text = ngq.PointRemark.ToString();
                                txtNGQty.Text = ngq.NGQty.ToString();
                            }


                            if (OpenPage > 0)
                                ShowImg();
                        }
                    }
                }else
                {
                    
                }
            }
            catch { }
        }
        private void btnAddNG_Click(object sender, EventArgs e)
        {
            if(MessageBox.Show("ต้องการบันทึก NG หรือไม่ ?","บันทึก",MessageBoxButtons.YesNo,MessageBoxIcon.Question)==DialogResult.Yes)
            {
                try
                {
                    if(!txtStatus.Text.Equals("Completed") && !txtNGID.Text.Equals("") && !txtNGQty.Text.Equals("") && !txtValue.Text.Equals(""))
                    {
                        int NGQ = 0;
                        int GQid = 0;
                        int.TryParse(txtNGID.Text, out GQid);
                        int.TryParse(txtNGQty.Text, out NGQ);
                        if (NGQ > 0 && GQid>0)
                        {
                            using (DataClasses1DataContext db = new DataClasses1DataContext())
                            {
                                tb_QCNGPoint qc = new tb_QCNGPoint();
                                if (!txtQCNo.Text.Equals(""))
                                {
                                    qc.QCNo = txtQCNo.Text;
                                }
                                else
                                {
                                    qc.QCNo = txtTempNo.Text;
                                }
                                qc.PTAG = PTAG;
                                qc.NCRNo = "";
                                qc.NCRSS = 0;
                                qc.NGQty = NGQ.ToString();
                                qc.ofTag = txtOfTAG.Text;
                                qc.PointRemark = txtValue.Text;
                                qc.PointRemark2 = "";                               
                                qc.SeqNo = 1;
                                qc.Status = "Waiting";
                                qc.WONo = txtProdNo.Text;
                                tb_QCGroupPart qg = db.tb_QCGroupParts.Where(q => q.FormISO.Equals(FormISO) && q.PartNo.Equals(txtPartNo.Text) && q.id.Equals(GQid)).FirstOrDefault();
                                if (qg != null)
                                {
                                    qc.PointName = qg.StepPart.ToString();
                                    qc.TopCaseText = qg.TopPic.ToString();
                                    qc.OldValue = qg.SetData.ToString();
                                    qc.SeqNo = qg.Seq;
                                }                             
                                
                                db.tb_QCNGPoints.InsertOnSubmit(qc);
                                db.SubmitChanges();
                               // cboCheckGroupPart.Text = "";
                               // txtValue.Text = "";
                              //  txtNGQty.Text = "";
                               // txtNGID.Text = "";
                                db.sp_46_QCHD_Update_HD(txtQCNo.Text.ToUpper()); //Call Update Status
                               // MessageBox.Show("บันทึกเรียบร้อย");
                                LoadNGPoint();
                            }
                        }
                    }
                }
                catch { }
            }
        }
        int rowNG = 0;
        private void radGridView2_CellClick(object sender, GridViewCellEventArgs e)
        {
            rowNG = e.RowIndex;
        }

        private void pictureBox1_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(globalPath);
            }
            catch { }
        }

        private void txtSeq_TextChanged(object sender, EventArgs e)
        {
           
        }


        string SEQOld = "";
        int Snew = 0;
        private void ShowImg()
        {
            try
            {
                txtValue1.Text = "";
                txtValue2.Text = "";
                txtValueCheck.Text = "";
                lblInvalidValue.Visible = false;
                chkPoint.Checked = false;
                chkPoint.Text = "ยังไม่ได้ตรวจสอบ";
                chkPoint.ForeColor = Color.Red;
                pictureBox1.Image = null;
                txtDataBox.Text = "";
                txtRemark.Text = "";
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    int seqs = 0;
                    int.TryParse(txtSeq.Text, out seqs);

                    tb_QCGroupPart im = db.tb_QCGroupParts.Where(u => u.PartNo.Equals(txtPartNo.Text) && u.Seq.Equals(seqs) && u.FormISO.Equals(FormISO)).FirstOrDefault();
                    
                    //if (ph != null)
                    //{
                    //    Path = ph.PathFile;
                    //}
                    if (im != null && !Path.Equals("") && im.Image1 !=null)
                    {
                        if (!im.Image1.ToString().Equals(""))
                        {
                            if (File.Exists(Path + im.Image1))
                            {
                                pictureBox1.Image = Image.FromFile(Path + im.Image1);
                                globalPath = Path + im.Image1;
                            }
                        }
                    }
                   
                    //check Value1
                    if (im != null)
                    {
                        txtValue1.Text = Convert.ToString(im.Value1);
                        txtValue2.Text = Convert.ToString(im.Value2);
                        txtValueCheck.Text = Convert.ToString(im.SetDate2);
                    }
                    lblInvalidValue.Visible = false;

                    tb_QCSetupPoint sp = db.tb_QCSetupPoints.Where(w => w.Seq.Equals(seqs) && w.FormISO.Equals(FormISO) && w.WONo.Equals(txtProdNo.Text) && w.PTAG.Equals(PTAG)).FirstOrDefault();
                    if (sp != null)
                    {
                        chkPoint.Checked = true;
                        chkPoint.ForeColor = Color.Green;
                        chkPoint.Text = "ตรวจสอบแล้ว";
                        txtDataBox.Text = Convert.ToString(sp.Value1);
                        txtRemark.Text = Convert.ToString(sp.Value5);
                        decimal Value3 = 0;
                        string SOK = "OK";
                        if (txtValueCheck.Text.Equals("Yes") && !txtDataBox.Text.Equals("") && decimal.TryParse(txtDataBox.Text, out Value3))
                        {
                            decimal Value1 = 0;
                            decimal Value2 = 0;
                           
                            decimal.TryParse(txtValue1.Text, out Value1);
                            decimal.TryParse(txtValue2.Text, out Value2);
                            decimal.TryParse(txtDataBox.Text, out Value3);
                            lblInvalidValue.Visible = false;
                            if (Value3<Value1)
                            {
                                lblInvalidValue.Visible = true;
                                lblInvalidValue.Text = "NG";
                                lblInvalidValue.ForeColor = Color.Red;
                                SOK = "NG";
                            }
                            if(Value3>Value2 && Value2>0)
                            {
                                lblInvalidValue.Visible = true;
                                lblInvalidValue.Text = "NG";
                                lblInvalidValue.ForeColor = Color.Red;
                                SOK = "NG";
                            }
                            if(SOK.Equals("OK"))
                            {
                                lblInvalidValue.Visible = true;
                                lblInvalidValue.Text = "OK";
                                lblInvalidValue.ForeColor = Color.Green;
                            }


                        }


                    }

                    if (seqs > 0)
                    {
                        tb_QCSetupPoint sp1 = db.tb_QCSetupPoints.Where(w => w.Seq.Equals(seqs) && w.FormISO.Equals(FormISO) && w.WONo.Equals(txtProdNo.Text) && w.PTAG.Equals(PTAG)).FirstOrDefault();
                        if (sp1 != null)
                        {
                            sp1.Value1 = txtDataBox.Text;
                            db.SubmitChanges();
                        }
                        else
                        {
                            if (FormISO.Equals("FM-PD-026_1") && seqs > 99)
                            {

                            }
                            if (FormISO.Equals("FM-PD-001") && seqs > 99)
                            {

                            }
                            if (FormISO.Equals("FM-PD-109") && seqs > 99)
                            {

                            }
                            if (FormISO.Equals("FM-PD-110") && seqs > 99)
                            {

                            }
                            if (FormISO.Equals("FM-PD-013") && seqs > 99)
                            {

                            }
                            if (FormISO.Equals("FM-PD-014") && seqs > 99)
                            {

                            }
                            else
                            {

                                tb_QCSetupPoint spn = new tb_QCSetupPoint();
                                spn.Seq = seqs;
                                spn.WONo = txtProdNo.Text;
                                spn.FormISO = FormISO;
                                spn.Value1 = txtDataBox.Text;
                                spn.Value2 = txtOfTAG.Text;
                                spn.Value3 = txtQCNo.Text;
                                spn.Value5 = txtRemark.Text;
                                spn.ValueDecimal = 0;
                                spn.ValueInt = 0;
                                spn.PTAG = PTAG;
                                spn.CheckBy = dbClss.UserID;
                                spn.CheckDate = DateTime.Now;
                                spn.Description = txtPartNo.Text;
                                db.tb_QCSetupPoints.InsertOnSubmit(spn);
                                db.SubmitChanges();
                            }
                        }
                    }


                    SEQOld = seqs.ToString();
                   

                }

            }
            catch(Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void LoadCheckValue()
        {
            try
            {

            }
            catch { }
        }

       

        private void radPageView1_PageIndexChanged(object sender, RadPageViewIndexChangedEventArgs e)
        {
           
        }

        private void radPageView1_SelectedPageChanged(object sender, EventArgs e)
        {
            Snew += 1;
            OpenPage = 1;
            txtDataBox.Focus();
        }

        private void txtDataBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if(e.KeyChar==13)
            {
                if (!txtDataBox.Text.Equals(""))
                {
                    try
                    {
                        int ids = 0;
                        string SOK = "";
                        int.TryParse(txtSeq.Text, out ids);
                        lblInvalidValue.Visible = false;
                        using (DataClasses1DataContext db = new DataClasses1DataContext())
                        {
                            tb_QCSetupPoint sp1 = db.tb_QCSetupPoints.Where(w => w.Seq.Equals(ids) && w.FormISO.Equals(FormISO) && w.WONo.Equals(txtProdNo.Text) && w.PTAG.Equals(PTAG)).FirstOrDefault();
                            if (sp1 != null)
                            {
                                sp1.Value1 = txtDataBox.Text;
                                db.SubmitChanges();

                                decimal Value3 = 0;

                                if (txtValueCheck.Text.Equals("Yes") && !txtDataBox.Text.Equals("") && decimal.TryParse(txtDataBox.Text, out Value3))
                                {
                                    SOK = "OK";
                                    decimal Value1 = 0;
                                    decimal Value2 = 0;
                                    decimal AAA = 0;
                                    decimal BBB = 0;
                                    decimal CCC = 0;

                                    decimal.TryParse(txtValue1.Text, out Value1);
                                    decimal.TryParse(txtValue2.Text, out Value2);
                                    decimal.TryParse(txtDataBox.Text, out Value3);
                                    lblInvalidValue.Visible = false;
                                    if (Value3 < Value1)
                                    {
                                        lblInvalidValue.Visible = true;
                                        lblInvalidValue.Text = "NG";
                                        lblInvalidValue.ForeColor = Color.Red;
                                        SOK = "NG";
                                    }
                                    if (Value3 > Value2 && Value2 > 0)
                                    {
                                        lblInvalidValue.Visible = true;
                                        lblInvalidValue.Text = "NG";
                                        lblInvalidValue.ForeColor = Color.Red;
                                        SOK = "NG";
                                    }
                                    if(Value1<0 || Value2<0)
                                    {
                                        AAA = 0;
                                       if(Value3>=Value1)
                                        {
                                            AAA = 1;
                                        }
                                       if(Value3<=Value2)
                                        {
                                            BBB = 1;
                                        }
                                       if(AAA==1 && BBB==1)
                                        {
                                            SOK = "OK";
                                        }else
                                        {
                                            lblInvalidValue.Visible = true;
                                            lblInvalidValue.Text = "NG";
                                            lblInvalidValue.ForeColor = Color.Red;
                                            SOK = "NG";
                                        }
                                    }

                                    if (SOK.Equals("OK"))
                                    {
                                        lblInvalidValue.Visible = true;
                                        lblInvalidValue.Text = "OK";
                                        lblInvalidValue.ForeColor = Color.Green;
                                    }

                                }
                                sp1.Value4 = SOK;
                                sp1.Value1 = txtDataBox.Text;
                                db.SubmitChanges();
                            }
                        }
                    }
                    catch { }
                }
                else
                {
                    using (DataClasses1DataContext db = new DataClasses1DataContext())
                    {
                        int ids = 0;
                        // string SOK = "";
                        int.TryParse(txtSeq.Text, out ids);
                        tb_QCSetupPoint sp1 = db.tb_QCSetupPoints.Where(w => w.Seq.Equals(ids) && w.FormISO.Equals(FormISO) && w.WONo.Equals(txtProdNo.Text) && w.PTAG.Equals(PTAG)).FirstOrDefault();
                        if (sp1 != null)
                        {
                            sp1.Value1 = txtDataBox.Text;
                            db.SubmitChanges();
                        }
                    }

                }
            }
        }

        private void radButton6_Click(object sender, EventArgs e)
        {
            if (!txtQCNo.Text.Equals(""))
            {
                QCListTAG lt = new QCListTAG(txtofTAGx,txtofTAGx2, txtQCNo.Text);
                lt.ShowDialog();
                if (!txtofTAGx.Text.Equals(""))
                {
                    txtDataBox.Text = "";
                    txtValue.Text = "";
                    txtNGQty.Text = "";
                    txtRemark.Text = "";
                    lblInvalidValue.Text = "";
                    txtNGPontQ.Text = "0";
                    txtNGInTAG.Text = "0";

                    PTAG = txtofTAGx.Text;
                    PTAG2 = txtofTAGx2.Text;
                    ClearGobalNo();
                    LoadData();
                }
            }
        }

        private void SelectNew()
        {
            ClearGobalNo();
            txtofTAGx.Text = "";
            if (FormISO.Equals("FM-PD-035_1") || FormISO.Equals("FM-QA-055_02_1") || FormISO.Equals("FM-QA-055")
                || FormISO.Equals("FM-PD-113") || FormISO.Equals("FM-PD-112") || FormISO.Equals("FM-PD-123") || FormISO.Equals("FM-PD-153") || FormISO.Equals("FM-PD-010"))
            {
                // PTAG = txtofTAGx.Text;
                txtScanTAG.Text = "";
                txtScanTAG.Focus();
                // LoadData();
                //
            }          
            else
            {
                if (!txtQCNo.Text.Equals(""))
                {
                    QCListTAG lt = new QCListTAG(txtofTAGx,txtofTAGx2, txtQCNo.Text);
                    lt.ShowDialog();
                    if (!txtofTAGx.Text.Equals(""))
                    {
                        PTAG = txtofTAGx.Text;
                        PTAG2 = txtofTAGx2.Text;
                        LoadData();
                    }
                }
            }
        }

        private void radButtonElement2_Click(object sender, EventArgs e)
        {
          
            SelectNew();
        }

        private void txtNGInTAG_KeyPress(object sender, KeyPressEventArgs e)
        {
            dbClss.CheckDigitDecimal(e);
            if(e.KeyChar==13)
            {
                try
                {
                    using (DataClasses1DataContext db = new DataClasses1DataContext())
                    {
                        int QTTNG = 0;
                        int.TryParse(txtNGInTAG.Text, out QTTNG);
                        tb_QCTAG qt = db.tb_QCTAGs.Where(q => q.BarcodeTag.Equals(PTAG) && q.QCNo.Equals(txtQCNo.Text)).FirstOrDefault();
                        if(qt!=null)
                        {
                            qt.NGofTAG = QTTNG;
                            db.SubmitChanges();
                            LoadNGPoint();
                            txtDataBox.Focus();
                        }
                    }
                }
                catch { }
            }
        }

        private void txtScanTAG_KeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {
                if (e.KeyChar == 13)
                {
                    if (!txtScanTAG.Text.Equals(""))
                    {
                        txtDataBox.Text = "";
                        txtValue.Text = "";
                        txtNGQty.Text = "";
                        txtRemark.Text = "";
                        lblInvalidValue.Text = "";
                        txtNGPontQ.Text = "0";
                        txtNGInTAG.Text = "0";
                        string[] DATA2 = txtScanTAG.Text.Split(',');
                        if (DATA2.Length == 8)
                        {
                            if (DATA2[7].Length >0)
                            {
                              
                                if (DATA2[1].Equals(WOs))
                                {
                                    PTAG = txtScanTAG.Text;
                                    PTAG2 = txtScanTAG.Text;
                                    using (DataClasses1DataContext db = new DataClasses1DataContext())
                                    {
                                        tb_QCTAG cktag = db.tb_QCTAGs.Where(t => t.BarcodeTag.Equals(PTAG) && t.QCNo.Equals(txtQCNo.Text)).FirstOrDefault();
                                        if (cktag == null)
                                        {
                                            tb_QCHD rlist = db.tb_QCHDs.Where(h => h.WONo.Equals(WOs) && h.FormISO.Equals(FormISO)).FirstOrDefault();
                                            if (rlist != null)
                                            {
                                                decimal dcq = 0;
                                                decimal.TryParse(DATA2[2], out dcq);
                                                if (DATA2.Length == 8)
                                                {
                                                    dbShowData.InsertTAG(PTAG, WOs, rlist.QCNo, dcq, DATA2[5], TypeP, rlist.LineName, "Normal", PTAG2);
                                                }
                                                ClearGobalNo();
                                                LoadData();
                                                // txtScanTAG.Text = "";
                                                // txtScanTAG.Focus();
                                            }
                                        }else
                                        {
                                            ClearGobalNo();
                                            LoadData();
                                        }
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Production No. ไม่ตรงกับรายการแรก!!!");
                                    txtScanTAG.Text = "";
                                    txtScanTAG.Focus();
                                }
                            }else
                            {
                                MessageBox.Show("1.TAG ไม่ถูกต้อง");
                            }
                        }
                        else
                        {
                            MessageBox.Show("2.TAG ไม่ถูกต้อง");
                        }
                    }
                }
            }
            catch { }
        }

        private void radButtonElement3_Click(object sender, EventArgs e)
        {
            if(!txtProdNo.Text.Equals("") && !txtQCNo.Text.Equals(""))
            {
                QCProblem qp = new QCProblem(txtQCNo.Text, txtProdNo.Text.ToUpper());
                qp.ShowDialog();
            }
        }

        private void txtInspector_KeyPress(object sender, KeyPressEventArgs e)
        {
            if(e.KeyChar==13)
            {
                if(!txtInspector.Text.Equals(""))
                {
                    dbShowData.InsertQCChecker(txtInspector.Text, txtQCNo.Text, "Ins", cboSelectCheckBy.Text,"");
                    txtInspector.Text = "";
                    LoadChecker();
                    txtInspector.Focus();
                }
            }
        }

        private void LoadChecker()
        {
            try
            {
                radGridView1.DataSource = null;
               // radGridView3.DataSource = null;
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    radGridView1.DataSource = db.tb_QCCheckUsers.Where(u => u.QCNo.Equals(txtQCNo.Text)).ToList();
                    //radGridView3.DataSource = db.tb_QCCheckUsers.Where(u => u.QCNo.Equals(txtQCNo.Text) && u.UType.Equals("Chk")).ToList();
                }

            }
            catch { }
        }

        private void txtChecker_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                if (!txtChecker.Text.Equals(""))
                {
                    dbShowData.InsertQCChecker(txtChecker.Text, txtQCNo.Text, "Chk",cboSelectCheckBy.Text,"");
                    txtChecker.Text = "";
                    LoadChecker();
                    txtChecker.Focus();
                }
            }
        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                int idC1 = 0;
                int.TryParse(radGridView1.Rows[rowC1].Cells["id"].Value.ToString(), out idC1);
                if (idC1 > 0)
                {
                    if (MessageBox.Show("ต้องการลบ ผู้ Inspector ?", "ลบข้อมูล", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        dbShowData.DeleteQCChecker(idC1);
                        LoadChecker();
                    }
                }
            }
            catch { }

        }

        private void deleteToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                int idC2 = 0;
                int.TryParse(radGridView3.Rows[rowC2].Cells["id"].Value.ToString(), out idC2);
                if (idC2 > 0)
                {
                    if (MessageBox.Show("ต้องการลบ ผู้ Check ?", "ลบข้อมูล", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        dbShowData.DeleteQCChecker(idC2);
                        LoadChecker();
                    }
                }
            }
            catch { }
        }
        int rowC1 = 0;
        private void radGridView1_CellClick_1(object sender, GridViewCellEventArgs e)
        {
            rowC1 = e.RowIndex;
        }
        int rowC2 = 0;
        private void radGridView3_CellClick(object sender, GridViewCellEventArgs e)
        {
            rowC2 = e.RowIndex;
        }

        private void txtInspector_Enter(object sender, EventArgs e)
        {
            dbShowData.Setcolor(txtInspector, 1);
        }

        private void txtInspector_Leave(object sender, EventArgs e)
        {
            dbShowData.Setcolor(txtInspector, 0);
        }

        private void txtNGInTAG_Enter(object sender, EventArgs e)
        {
            dbShowData.SetColor2(txtNGInTAG, 1);
        }

        private void txtNGInTAG_Leave(object sender, EventArgs e)
        {
            dbShowData.SetColor2(txtNGInTAG, 0);
        }

        private void txtChecker_Leave(object sender, EventArgs e)
        {
            dbShowData.Setcolor(txtChecker, 0);
        }

        private void txtChecker_Enter(object sender, EventArgs e)
        {
            dbShowData.Setcolor(txtChecker, 1);
            
        }

        private void txtValue_Enter(object sender, EventArgs e)
        {
            dbShowData.SetColor2(txtValue, 1);
        }

        private void txtValue_Leave(object sender, EventArgs e)
        {
            dbShowData.SetColor2(txtValue, 0);
        }

        private void txtNGQty_Enter(object sender, EventArgs e)
        {
            dbShowData.SetColor2(txtNGQty, 1);
        }

        private void txtNGQty_Leave(object sender, EventArgs e)
        {
            dbShowData.SetColor2(txtNGQty, 0);
        }

        private void radButton7_Click(object sender, EventArgs e)
        {
            //List NG
            if(!PTAG.Equals(""))
            {
                QCListNGTAG ngt = new QCListNGTAG(txtQCNo.Text,PTAG);
                ngt.Show();
            }
        }

        private void btnUP_Click(object sender, EventArgs e)
        {
            UpDownData(0);
        }

        private void btnDown_Click(object sender, EventArgs e)
        {
            UpDownData(1);
        }

        private void txtRemark_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                if (!txtRemark.Text.Equals(""))
                {
                    try
                    {
                        int ids = 0;
                       
                        int.TryParse(txtSeq.Text, out ids);
                        lblInvalidValue.Visible = false;
                        using (DataClasses1DataContext db = new DataClasses1DataContext())
                        {
                           
                                tb_QCSetupPoint sp1 = db.tb_QCSetupPoints.Where(w => w.Seq.Equals(ids) && w.FormISO.Equals(FormISO) && w.WONo.Equals(txtProdNo.Text) && w.PTAG.Equals(PTAG)).FirstOrDefault();
                                if (sp1 != null)
                                {
                                    sp1.Value5 = txtRemark.Text;
                                    db.SubmitChanges();
                                }
                           
                        }
                    }
                    catch { }
                }
            }
        }

        private void radGridView1_CellEndEdit_1(object sender, GridViewCellEventArgs e)
        {
            try
            {
                if (radGridView1.Columns["BoxNo"].Index==e.ColumnIndex && e.RowIndex>=0)
                {
                    using (DataClasses1DataContext db = new DataClasses1DataContext())
                    {
                        tb_QCCheckUser qc = db.tb_QCCheckUsers.Where(p => p.id.Equals(Convert.ToInt32(radGridView1.CurrentRow.Cells["id"].Value.ToString()))).FirstOrDefault();
                        if(qc!=null)
                        {
                            qc.BoxNo = radGridView1.CurrentRow.Cells["BoxNo"].Value.ToString();
                            db.SubmitChanges();
                        }
                    }
                }
            }
            catch { }
        }

        private void radButtonElement4_Click(object sender, EventArgs e)
        {
            QCCheckUserPoint2 qp2 = new QCCheckUserPoint2(txtQCNo.Text);
            qp2.ShowDialog();
        }

        private void radButtonElement5_Click(object sender, EventArgs e)
        {
            dbShowData.ExportScanTAG(txtProdNo.Text);
        }
    }
}
