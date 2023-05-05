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
    public static class dbShowData
    {
        public static void Setcolor(TextBox tx,int AC)
        {
            if (AC.Equals(1))
            {
                tx.BackColor = Color.Yellow;
            }
            else
            {
                tx.BackColor = Color.White;
            }

        }
        public static void SetColor2(RadTextBox tx,int AC)
        {
            if (AC.Equals(1))
            {
                tx.BackColor = Color.Yellow;
            }
            else
            {
                tx.BackColor = Color.White;
            }
        }
        public static void CreateListQC(string WO,string ISO)
        {
            try
            {
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    string QCNo = "";
                    tb_QCHD qch = db.tb_QCHDs.Where(q => q.WONo.Equals(WO) && q.FormISO.Equals(ISO)).FirstOrDefault();
                    if (qch == null)
                    {
                        var lWo = db.sp_46_QCSelectWO_01(WO).FirstOrDefault();
                        if (lWo != null)
                        {
                            QCNo = dbClss.GetSeriesNo(6, 2);
                            tb_QCHD qcN = new tb_QCHD();
                            qcN.CheckBy1 = "";
                            qcN.CheckBy2 = "";
                            qcN.CheckBy3 = "";
                            qcN.IssueBy = "";
                            qcN.IssueBy2 = "";
                            qcN.ApproveBy = "";
                            qcN.ApproveBy2 = "";                           
                            qcN.QCNo = QCNo;
                            qcN.WONo = WO;
                            qcN.PartNo = lWo.CODE;
                            qcN.OrderQty = lWo.OrderQty;
                            qcN.OKQty = 0;
                            qcN.NGQty = 0;
                            qcN.LotNo = lWo.LotNo;
                            qcN.LineName = lWo.BUMO;
                            qcN.CreateBy = dbClss.UserID;
                            qcN.CreateDate = DateTime.Now;
                            qcN.SS = 1;
                            qcN.Status = "Checking";
                            qcN.SendApprove = false;
                            qcN.FormISO = ISO;
                            qcN.DocRef1 = "";
                            qcN.DocRef2 = "";
                            qcN.ApproveBy = "";
                            qcN.ApproveBy2 = "";
                            qcN.CheckBy1 = "";
                            qcN.CheckBy2 = "";
                            qcN.CheckBy3 = "";
                            
                            qcN.IssueBy = "";
                            qcN.IssueBy2 = "";
                            qcN.ChangeModel = lWo.ChangeModel;
                            qcN.DayNight = lWo.DayNight;
                            qcN.QCPoint = "";
                            qcN.RefValue1 = "";
                            qcN.RefValue2 = "";
                            qcN.RefValue3 = "";
                            db.tb_QCHDs.InsertOnSubmit(qcN);
                            db.SubmitChanges();
                        }
                    }
                }
            }
            catch { }
        }
        public static void InsertTAG(string PTAG,string WO,string QCNo,decimal Qty,string ofTAG,string Type,string LineNo,string GType,string TAG)
        {
            try
            {
                int OK = 0;
                int NG = 0;               
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {

                    //SumQty Inspection / NG / OK 
                    decimal SumALL = Convert.ToDecimal(db.get_QCSumQty(QCNo, PTAG, 5));
                    decimal SumOK = 0;// Convert.ToDecimal(db.get_QCSumQty(Upd.QCNo, 1));
                    decimal SumNG = Convert.ToDecimal(db.get_QCSumQty(QCNo, PTAG, 4));
                    SumOK = SumALL - SumNG;


                    tb_QCTAG qctag = db.tb_QCTAGs.Where(t => t.BarcodeTag.Equals(PTAG) && t.QCNo.Equals(QCNo)).FirstOrDefault();
                    if (qctag == null)
                    {
                        tb_QCTAG qct = new tb_QCTAG();
                        qct.QCNo = QCNo;
                        qct.BarcodeTag = PTAG;
                        qct.SS = 1;
                        qct.QtyofTag = Qty;
                        qct.OKQty = (Qty - NG);
                        qct.NGQty = NG;
                        qct.ofTAG = ofTAG;
                        qct.Dept = Type;
                        qct.CheckDate = DateTime.Now;
                        qct.CheckBy = dbClss.UserID;
                        qct.DType = LineNo;
                        qct.NGofTAG = 0;
                        qct.Seq = 1;
                        qct.CheckTAG = false;
                        qct.GTAG = TAG;
                        qct.GType = GType;
                        db.tb_QCTAGs.InsertOnSubmit(qct);
                        db.SubmitChanges();
                    }
                    //else
                    //{

                    //    qctag.QtyofTag = Qty;
                    //    qctag.OKQty = (OK - NG);
                    //    qctag.NGQty = NG;
                    //    qctag.ofTAG = ofTAG;
                    //    db.SubmitChanges();
                    //}
                }
            }
            catch { }
        }
        public static void InsertScanMachine(string WO,string TAG,string FormISO,string PartNo)
        {
            try
            {
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    string[] DATA = TAG.Split(',');
                    if(DATA.Length==1)
                    {
                        int Seqa = 0;
                        string DN = dbShowData.CheckDayN(DateTime.Now);
                        string MC = TAG;// DATA[0];
                       // DN = "N";
                       // int.TryParse(DATA[0], out Seqa);
                        if (MC != "")
                        {
                           // db.sp_46_QCMachine_Copy(FormISO, PartNo, WO);
                            tb_QCCheckMachine qc = db.tb_QCCheckMachines.Where(q => q.WONo.Equals(WO) && q.FormISO.Equals(FormISO) && q.UMC.Equals(MC) && q.DayN.Equals("")).FirstOrDefault();
                            if (qc != null)
                            {
                                qc.SC = "OK";
                                qc.CreateBy = dbClss.UserID;
                                qc.CreateDate = DateTime.Now;
                                qc.TAGScan = TAG;
                                qc.DayN = DN;
                                db.SubmitChanges();
                            }                          
                            else
                            {
                                tb_QCCheckMachine qc2 = db.tb_QCCheckMachines.Where(q => q.WONo.Equals(WO) && q.FormISO.Equals(FormISO) && q.UMC.Equals(MC) && q.DayN.Equals(DN)).FirstOrDefault();
                                if (qc2 != null)
                                {
                                    qc2.SC = "OK";
                                    qc2.CreateBy = dbClss.UserID;
                                    qc2.CreateDate = DateTime.Now;
                                    qc2.TAGScan = TAG;
                                    qc2.DayN = DN;
                                    db.SubmitChanges();
                                }
                                else
                                {
                                    tb_QCCheckMachine qc3 = db.tb_QCCheckMachines.Where(q => q.WONo.Equals(WO) && q.FormISO.Equals(FormISO) && q.UMC.Equals(MC) && q.DayN.Equals(DN)).FirstOrDefault();
                                    if (qc3 == null)
                                    {
                                        tb_QCGroupPart qcg = db.tb_QCGroupParts.Where(p => p.FormISO.Equals(FormISO) && p.PartNo.Equals(PartNo) && p.UseMachine.Equals(MC)).FirstOrDefault();
                                        if(qcg!=null)
                                        {
                                            Seqa = Convert.ToInt32(qcg.Seq);
                                        }

                                        tb_QCCheckMachine qn = new tb_QCCheckMachine();
                                        qn.WONo = WO;
                                        qn.FormISO = FormISO;
                                        qn.Seq = Seqa;
                                        qn.TAGScan = TAG;
                                        qn.PartNo = PartNo;
                                        qn.CreateBy = dbClss.UserID;
                                        qn.CreateDate = DateTime.Now;
                                        qn.DayN = DN;
                                        qn.SC = "OK";
                                        qn.UMC = MC;
                                        db.tb_QCCheckMachines.InsertOnSubmit(qn);
                                        db.SubmitChanges();
                                    }
                                }
                            }
                            db.sp_46_QCUpdate_Machine_Seq(WO, PartNo, FormISO);
                        }
                    }
                }
            }
            catch(Exception ex) { MessageBox.Show(ex.Message); }
        }
        public static string GETQCNo(string WO,string ISO)
        {
            string QCNo = "";
            using (DataClasses1DataContext db = new DataClasses1DataContext())
            {
                tb_QCHD qh = db.tb_QCHDs.Where(q => q.FormISO.Equals(ISO) && q.WONo.Equals(WO)).FirstOrDefault();
                if (qh != null)
                {
                    QCNo = qh.QCNo;
                }
            }

                return QCNo;
        }
        public static string CheckDayN(DateTime date1)
        {
            string DayN = "D";
            
            try
            {
               // date1 = Convert.ToDateTime("08:10:00") - 17:10
                TimeSpan ts = new TimeSpan(date1.Hour, date1.Minute, date1.Second);
                if (ts.TotalMinutes >= 495 && ts.TotalMinutes < 1215)
                //if (ts.TotalMinutes >= 510 && ts.TotalMinutes < 650)
                {
                    DayN = "D";
                }
                else
                {
                    DayN = "N";
                }              
            }
            catch { }
            return DayN;
        }
        public static int CheckColorDayN(string PartNo,string WONo)
        {
            int AC = 0;
            try
            {
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    int A1 = 0;
                    int A2 = 0;
                    var cklist = db.tb_QCCheckParts.Where(c => c.PartNo.Equals(PartNo) && c.OrderNo.Equals(WONo)).ToList();
                    foreach(var rd in cklist)
                    {
                        if(rd.DayN.Equals("D"))
                        {
                            A1 += 1;
                        }
                        if(rd.DayN.Equals("N"))
                        {
                            A2 += 1;
                        }
                    }

                    if (A1 > 0 && A2 > 0)
                    {
                        AC = 3;
                    }
                    else if (A1 > 0 && A2 == 0)
                    {
                        AC = 1;
                    }
                    else if (A1 == 0 && A2 > 0)
                    {
                        AC = 2;
                    }


                }
            }
            catch { }
            return AC;
        }
        public static string FormISOName(string ISO, DateTime Datex)
        {
            string RT = "";

            return RT;
        }
        public static void Print026A01012020(string WO, string PartNo, string QCNo1,string FileName)
        {
            try
            {

                string DATA = AppDomain.CurrentDomain.BaseDirectory;
                string tempPath = System.IO.Path.GetTempPath();
                //string FileName = "FM-PD-026.xlsx";
                ////FM-PD-026_17Feb23.xlsx
                //FileName = "FM-PD-026.07.2021.xlsx";
                //using (DataClasses1DataContext db = new DataClasses1DataContext())
                //{
                //    tb_ProductionHD pd1 = db.tb_ProductionHDs.Where(p => p.OrderNo.Equals(WO)).FirstOrDefault();
                //    if (pd1 != null)
                //    {
                //        DateTime Date1 = Convert.ToDateTime("2023-02-17 00:00:00.000");
                //        DateTime Date2 = Convert.ToDateTime("2023-03-01 00:00:00.000");

                //        if (pd1.Createdate >= Date1)
                //        {
                //            FileName = "FM-PD-026_17Feb23.xlsx";
                //        }
                //        if (pd1.Createdate >= Date2)
                //        {
                //            FileName = "FM-PD-026.01Mar23.xlsx";
                //        }

                //    }
                //}



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
                //int row1 = 22;
                //int row2 = 22;
                //int Seq = 0;
                //int seq2 = 22;
                //int CountRow = 0;
                string cIssueBy1 = "";
                string cIssueBy2 = "";
                string cIssueBy3 = "";
                string cIssueBy4 = "";

                string cCheckBy1 = "";
                string cCheckBy2 = "";
                string cCheckBy3 = "";

                string cCheckByF1 = "";
                string cCheckByF2 = "";
                string cCheckByF3 = "";

                string PV = "P";
                string QHNo = QCNo1;
                string FormISO = "";
                string DN = "";
                string SymBo = "～";
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    //string Value1 = "";
                    //string Value2 = "";
                    //string LotNo = "";



                    ///////////////SETValue/////////////////
                    var DValue = db.sp_46_QCSelectWO_01(WO).FirstOrDefault();
                    if (DValue != null)
                    {
                        DN = DValue.DayNight;
                        Excel.Range CPart = worksheet.get_Range("P3");
                        CPart.Value2 = DValue.CODE;

                        Excel.Range CStamp = worksheet.get_Range("X5");
                        if (PartNo.Length > 0)
                        {
                            if (dbClss.Right(PartNo, 1).Equals("W"))
                            {
                                CStamp.Value2 = "'" + dbClss.Right(PartNo, 8).Substring(0, 2) + "  " + dbClss.Right(PartNo, 6).Substring(0, 5);
                            }
                            else
                            {
                                CStamp.Value2 = "'" + dbClss.Right(PartNo, 7).Substring(0, 2) + "  " + dbClss.Right(PartNo, 5);
                            }
                        }


                        Excel.Range CName = worksheet.get_Range("I5");
                        CName.Value2 = DValue.NAME;

                        Excel.Range WOs = worksheet.get_Range("D5");
                        WOs.Value2 = DValue.PORDER;

                        Excel.Range CDate = worksheet.get_Range("D7");
                        CDate.Value2 = DValue.DeliveryDate;

                        Excel.Range CLot = worksheet.get_Range("D9");
                        CLot.Value2 = DValue.LotNo;

                        Excel.Range CQty = worksheet.get_Range("D11");
                        CQty.Value2 = DValue.OrderQty.ToString();


                        try
                        {
                            tb_QCHD qh = db.tb_QCHDs.Where(w => w.QCNo.Equals(QCNo1)).FirstOrDefault();
                            if (qh != null)
                            {
                                //////////Find UserName////////////
                                var uc = db.tb_QCCheckUsers.Where(u => u.QCNo.Equals(QCNo1)).ToList();
                                int r1 = 0;
                                int r2 = 0;
                                int r3 = 0;
                                int rr1 = 0;
                                int rr2 = 0;
                                int rr3 = 0;

                                foreach (var rd in uc)
                                {
                                    DN = rd.DayN;// dbShowData.CheckDayN(Convert.ToDateTime(rd.ScanDate));



                                    if (DN.Equals("D"))
                                    {
                                        if (rd.UDesc.Equals("ผู้จัดทำเอกสาร"))
                                            cIssueBy1 = rd.UserName;
                                        if (rd.UDesc.Equals("ผู้ตรวจสอบก่อนผลิต"))
                                            cIssueBy2 = rd.UserName;

                                        if (rd.UDesc.Equals("พนักงานประกอบ SUB LINE"))
                                        {
                                            r1 += 1;
                                            if (r1 > 1)
                                                cCheckBy1 += ",";
                                            cCheckBy1 += rd.UserName;
                                        }
                                        else if (rd.UDesc.Equals("พนักงานประกอบ MAIN LINE"))
                                        {
                                            r2 += 1;
                                            if (r2 > 1)
                                                cCheckBy2 += ",";
                                            cCheckBy2 += rd.UserName;
                                        }
                                        else if (rd.UDesc.Equals("พนักงานประกอบ FINAL LINE"))
                                        {
                                            r3 += 1;
                                            if (r3 > 1)
                                                cCheckBy3 += ",";
                                            cCheckBy3 += rd.UserName;
                                        }
                                    }
                                    else //N
                                    {
                                        if (rd.UDesc.Equals("ผู้จัดทำเอกสาร"))
                                            cIssueBy3 = rd.UserName;
                                        if (rd.UDesc.Equals("ผู้ตรวจสอบก่อนผลิต"))
                                            cIssueBy4 = rd.UserName;

                                        if (rd.UDesc.Equals("พนักงานประกอบ SUB LINE"))
                                        {
                                            rr1 += 1;
                                            if (rr1 > 1)
                                                cCheckByF1 += ",";
                                            cCheckByF1 += rd.UserName;


                                        }
                                        else if (rd.UDesc.Equals("พนักงานประกอบ MAIN LINE"))
                                        {
                                            rr2 += 1;
                                            if (rr2 > 1)
                                                cCheckByF2 += ",";
                                            cCheckByF2 += rd.UserName;

                                        }
                                        else if (rd.UDesc.Equals("พนักงานประกอบ FINAL LINE"))
                                        {
                                            rr3 += 1;
                                            if (rr3 > 1)
                                                cCheckByF3 += ",";
                                            cCheckByF3 += rd.UserName;

                                        }
                                    }
                                }


                                FormISO = qh.FormISO;
                                QHNo = qh.QCNo;
                                Excel.Range Ap = worksheet.get_Range("AE10");
                                Ap.Value2 = db.QC_GetUserName(qh.ApproveBy);// Convert.ToString(qh.ApproveBy);


                                Excel.Range CheckBy1 = worksheet.get_Range("E23");
                                CheckBy1.Value2 = cCheckBy1;
                                Excel.Range CheckBy2 = worksheet.get_Range("E32");
                                CheckBy2.Value2 = cCheckBy2;
                                Excel.Range CheckBy3 = worksheet.get_Range("E40");
                                CheckBy3.Value2 = cCheckBy3;

                                Excel.Range CheckByF1 = worksheet.get_Range("F23");
                                CheckByF1.Value2 = cCheckByF1;
                                Excel.Range CheckByF2 = worksheet.get_Range("F32");
                                CheckByF2.Value2 = cCheckByF2;
                                Excel.Range CheckByF3 = worksheet.get_Range("F40");
                                CheckByF3.Value2 = cCheckByF3;


                                //if (DN.Equals("D"))
                                //{
                                Excel.Range IssueBy = worksheet.get_Range("AE5");
                                IssueBy.Value2 = "1. " + cIssueBy1;
                                Excel.Range IssueBy2 = worksheet.get_Range("AE7");
                                IssueBy2.Value2 = "2. " + cIssueBy2;
                                //}
                                //else
                                //{
                                Excel.Range IssueBy3 = worksheet.get_Range("AF5");
                                IssueBy3.Value2 = "1. " + cIssueBy3;
                                Excel.Range IssueBy4 = worksheet.get_Range("AF7");
                                IssueBy4.Value2 = "2. " + cIssueBy4;
                                //}

                                QHNo = qh.QCNo;

                                ////Set Topic//

                                Excel.Range AF1 = worksheet.get_Range("AF1");
                                AF1.Value2 = "'" + db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 56);

                                //Excel.Range AA40 = worksheet.get_Range("AA40");
                                //AA40.Value2 = db.get_QC_SetDataMasterTpic(qh.FormISO, qh.PartNo, 38);

                                //Excel.Range AA32 = worksheet.get_Range("AA32");
                                //AA32.Value2 = db.get_QC_SetDataMasterTpic(qh.FormISO, qh.PartNo, 30);

                                //Step 1
                                int cRow = 41;
                                string Ppart = "";
                                for (int II = 20; II <= 22; II++)
                                {
                                    cRow += 1;

                                    ////Line 1 //
                                    Ppart = db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, II);

                                    Excel.Range Line1 = worksheet.get_Range("L" + cRow.ToString());
                                    Line1.Value2 = Ppart;

                                    var rds = db.sp_46_QCGetValue2601(qh.WONo, Ppart).FirstOrDefault();
                                    if (rds != null)
                                    {

                                        Excel.Range Line2 = worksheet.get_Range("Q" + cRow.ToString());
                                        Line2.Value2 = rds.DayN;

                                        Excel.Range Line3 = worksheet.get_Range("R" + cRow.ToString());
                                        Line3.Value2 = rds.NightN;

                                        Excel.Range Line4 = worksheet.get_Range("S" + cRow.ToString());
                                        Line4.Value2 = "'" + rds.Lot;
                                    }

                                }
                                /////FIND Part Name,Part No.
                                string Pparg2 = "";
                                //var QCP = db.sp_46_QCSelectWO_13_GroupPart(qh.WONo).ToList();
                                var QCP = db.tb_ProductionRMs.Where(p => p.OrderNo.Equals(qh.WONo)).ToList();
                                cRow = 22;
                                foreach (var rx in QCP)
                                {
                                    Ppart = "";
                                    Pparg2 = "";

                                    cRow += 1;
                                    if (cRow < 42)
                                    {

                                        Ppart = rx.PartNoRM;
                                        Pparg2 = db.getItemNo(rx.PartNoRM);
                                        if (rx.PartNoRM.ToUpper().Contains("GREASE") || rx.PartNoRM.ToUpper().Contains("LOCTITE"))
                                        {
                                            //Pparg2 = rx.PartNo;
                                            //if (rx.PartNo.ToUpper().Contains("GREASE"))
                                            //    Ppart = db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 20);
                                            //if (rx.PartNo.ToUpper().Contains("LOCTITE 277"))
                                            //    Ppart = db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 21);
                                            //if (rx.PartNo.ToUpper().Contains("LOCTITE 414"))
                                            //    Ppart = db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 22);

                                            Pparg2 = "";
                                        }
                                        if (Pparg2 != "")
                                        {
                                            InsertToExcel(ref worksheet, "G" + cRow.ToString(), Pparg2);
                                            InsertToExcel(ref worksheet, "L" + cRow.ToString(), Ppart);
                                            Ppart = rx.PartNoRM;
                                            var rds = db.sp_46_QCGetValue2601(qh.WONo, Ppart).FirstOrDefault();
                                            if (rds != null)
                                            {
                                                InsertToExcel(ref worksheet, "Q" + cRow.ToString(), rds.DayN);
                                                InsertToExcel(ref worksheet, "R" + cRow.ToString(), rds.NightN);
                                                InsertToExcel(ref worksheet, "S" + cRow.ToString(), rds.Lot);
                                            }
                                        }
                                    }
                                }


                                //Step 2
                                int crow2 = 22;
                                cRow = 22;
                                int NewR = 0;
                                int NewR2 = 0;
                                string CheckValueSetup = "";
                                int A35 = 0;
                                int A36 = 0;
                                int A37 = 0;
                                int A38 = 0;
                                int D23 = 0;
                                int N23 = 0;
                                for (int II = 23; II <= 55; II++)
                                {
                                    cRow += 1;
                                    crow2 += 1;
                                    CheckValueSetup = "";
                                    ////Line 1 //
                                    if (II != 29)
                                    {
                                        //การเซ็ต
                                        NewR2 = II;
                                        NewR = cRow;
                                        //if (II >= 47)
                                        //{
                                        //    NewR2 = II + 5;
                                        //    if (II >= 52)
                                        //    {
                                        //        NewR2 = II - 5;
                                        //        NewR = NewR - 1;
                                        //    }

                                        //}

                                        Excel.Range Line1 = worksheet.get_Range("AE" + NewR.ToString());
                                        CheckValueSetup = db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, II);
                                        Line1.Value2 = CheckValueSetup;




                                        if (II == 48 || II == 49)
                                        {
                                            Excel.Range Line2 = worksheet.get_Range("AE" + NewR.ToString());
                                            Line1.Value2 = db.QC_GetP26LotShift(qh.WONo, II);
                                        }

                                    }



                                    if (II != 42)
                                    {

                                        NewR2 = II;
                                        NewR = crow2;


                                        var rss = db.sp_46_QCGetValue2601_20(qh.WONo, II).FirstOrDefault();
                                        if (rss != null)
                                        {
                                            if (!rss.DayN.Equals(""))
                                            {
                                                Excel.Range Line2 = worksheet.get_Range("AG" + II.ToString());
                                                Line2.Value2 = rss.DayN;
                                                D23 = 1;
                                            }
                                            if (!rss.NightN.Equals(""))
                                            {
                                                Excel.Range Line3 = worksheet.get_Range("AH" + II.ToString());
                                                Line3.Value2 = rss.NightN;
                                                N23 = 1;
                                            }
                                            //Skip1//////////////////////////////////////
                                            if (CheckValueSetup.Contains("ไม่มี"))
                                            {
                                                if (D23 > 0)
                                                {
                                                    Excel.Range LineVS1 = worksheet.get_Range("AG" + II.ToString());
                                                    LineVS1.Value2 = "P";
                                                }
                                                if (N23 > 0)
                                                {
                                                    Excel.Range LineVS2 = worksheet.get_Range("AH" + II.ToString());
                                                    LineVS2.Value2 = "P";
                                                }
                                            }
                                            //Skip2//////////////////////////////////////
                                            if (II == 35)
                                            {
                                                if (rss.DayN.Equals("P"))
                                                {
                                                    A35 = 1;
                                                }
                                                if (rss.NightN.Equals("P"))
                                                {
                                                    A36 = 1;
                                                }
                                            }
                                            if (II == 38)
                                            {
                                                if (rss.DayN.Equals("P"))
                                                {
                                                    A37 = 1;
                                                }
                                                if (rss.NightN.Equals("P"))
                                                {
                                                    A38 = 1;
                                                }
                                            }
                                            if (II == 36)
                                            {
                                                if (A35 == 1)
                                                {
                                                    Excel.Range Line36 = worksheet.get_Range("AG" + II.ToString());
                                                    Line36.Value2 = "P";
                                                }
                                                if (A36 == 1)
                                                {
                                                    Excel.Range Line36 = worksheet.get_Range("AH" + II.ToString());
                                                    Line36.Value2 = "P";
                                                }
                                            }
                                            if (II == 39)
                                            {
                                                if (A37 == 1)
                                                {
                                                    Excel.Range Line39 = worksheet.get_Range("AG" + II.ToString());
                                                    Line39.Value2 = "P";
                                                }
                                                if (A38 == 1)
                                                {
                                                    Excel.Range Line39 = worksheet.get_Range("AH" + II.ToString());
                                                    Line39.Value2 = "P";
                                                }
                                            }

                                            /////////////////////////////////////////////

                                        }
                                    }
                                    else
                                    {
                                        Excel.Range Line2 = worksheet.get_Range("AG42");
                                        Line2.Value2 = db.get_QC_DATAPoint_AG(qh.WONo, 42);
                                    }



                                }

                                Excel.Range Loctite1 = worksheet.get_Range("S42");
                                Loctite1.Value2 = "'" + db.get_QC_ValueRM(WO, "GREASE G-30M", 20);
                                Excel.Range Loctite2 = worksheet.get_Range("S43");
                                Loctite2.Value2 = "'" + db.get_QC_ValueRM(WO, "LOCTITE 277", 21);
                                Excel.Range Loctite3 = worksheet.get_Range("S44");
                                Loctite3.Value2 = "'" + db.get_QC_ValueRM(WO, "LOCTITE 414", 22);

                                string DDN1 = db.get_QC_ValueRM22(WO, "GREASE G-30M", 20);
                                string DDN2 = db.get_QC_ValueRM22(WO, "LOCTITE 277", 21);
                                string DDN3 = db.get_QC_ValueRM22(WO, "LOCTITE 414", 22);

                                //Step 1
                                if (DDN1.Equals("D") || DDN1.Equals("A"))
                                {
                                    Excel.Range LoctiteQ1 = worksheet.get_Range("Q42");
                                    LoctiteQ1.Value2 = "P";
                                }
                                if (DDN1.Equals("N") || DDN1.Equals("A"))
                                {
                                    Excel.Range LoctiteR1 = worksheet.get_Range("R42");
                                    LoctiteR1.Value2 = "P";
                                }

                                //Step 2
                                if (DDN2.Equals("D") || DDN2.Equals("A"))
                                {
                                    Excel.Range LoctiteQ2 = worksheet.get_Range("Q43");
                                    LoctiteQ2.Value2 = "P";
                                }
                                if (DDN2.Equals("N") || DDN2.Equals("A"))
                                {
                                    Excel.Range LoctiteR2 = worksheet.get_Range("R43");
                                    LoctiteR2.Value2 = "P";
                                }

                                //Step 3
                                if (DDN3.Equals("D") || DDN3.Equals("A"))
                                {
                                    Excel.Range LoctiteQ3 = worksheet.get_Range("Q44");
                                    LoctiteQ3.Value2 = "P";
                                }
                                if (DDN3.Equals("N") || DDN3.Equals("A"))
                                {
                                    Excel.Range LoctiteR3 = worksheet.get_Range("R44");
                                    LoctiteR3.Value2 = "P";
                                }







                            }
                            var gTime = db.sp_46_QCGetValue2601_Time(WO).ToList();
                            if (gTime.Count > 0)
                            {
                                var g = gTime.FirstOrDefault();
                                DateTime Chtime = Convert.ToDateTime(g.BomTime);
                                DateTime Chtime2 = Convert.ToDateTime(g.PrintTime);
                                if (g.BomTime == g.PrintTime)
                                {
                                    Chtime2 = Convert.ToDateTime(g.PrintTime).AddMinutes(30);
                                }

                                Excel.Range AB = worksheet.get_Range("AB9");
                                AB.Value2 = Math.Abs(Convert.ToDecimal((Chtime - Chtime2).TotalMinutes)).ToString("####") + " นาที";

                                if (!g.StartTime.Equals(""))
                                {
                                    Excel.Range StartT = worksheet.get_Range("N7");
                                    StartT.Value2 = Convert.ToDateTime(Chtime2).ToString("HH:mm");

                                    Excel.Range EndT = worksheet.get_Range("AA7");
                                    EndT.Value2 = Convert.ToDateTime(g.EndTime).ToString("HH:mm");

                                    // int ChanP = 0;
                                    //int.TryParse(Convert.ToInt32(DValue.ChangeModel).ToString(), out ChanP);
                                    // if (ChanP > 0)
                                    // {

                                    Excel.Range O9 = worksheet.get_Range("O9");
                                    O9.Value2 = "'" + Convert.ToDateTime(g.BomTime).ToString("HH:mm") + "-" + Convert.ToDateTime(Chtime2).ToString("HH:mm");

                                    //}

                                }
                            }

                            //Find Problem//

                            tb_QCProblem pb = db.tb_QCProblems.Where(p => p.QCNo.Equals(QHNo)).FirstOrDefault();
                            if (pb != null)
                            {
                                if (pb.TypeProblem.Equals("Man"))
                                {
                                    Excel.Range PBA = worksheet.get_Range("F13");
                                    PBA.Value2 = "P";

                                }
                                else if (pb.TypeProblem.Equals("Machine"))
                                {
                                    Excel.Range PBA = worksheet.get_Range("I13");
                                    PBA.Value2 = "P";
                                }
                                else if (pb.TypeProblem.Equals("Method"))
                                {
                                    Excel.Range PBA = worksheet.get_Range("M13");
                                    PBA.Value2 = "P";
                                }
                                else if (pb.TypeProblem.Equals("Material"))
                                {
                                    Excel.Range PBA = worksheet.get_Range("P13");
                                    PBA.Value2 = "P";
                                }
                                else if (pb.TypeProblem.Equals("Other"))
                                {
                                    Excel.Range PBA = worksheet.get_Range("S13");
                                    PBA.Value2 = "P";
                                    Excel.Range PBA2 = worksheet.get_Range("X13");
                                    PBA2.Value2 = pb.TypeRemark;
                                }

                                Excel.Range PC1 = worksheet.get_Range("F14");
                                PC1.Value2 = pb.ProblemSeeBy;
                                Excel.Range PC2 = worksheet.get_Range("N14");
                                PC2.Value2 = pb.ProblemName;

                                Excel.Range PC3 = worksheet.get_Range("AC14");
                                PC3.Value2 = pb.ProblemWare;
                                Excel.Range PC4 = worksheet.get_Range("F15");
                                PC4.Value2 = pb.ProblemTime;

                                Excel.Range PC5 = worksheet.get_Range("N15");
                                PC5.Value2 = pb.ProblemWhy;

                                Excel.Range PC6 = worksheet.get_Range("G17");
                                PC6.Value2 = pb.ProblemFix;
                                Excel.Range PC7 = worksheet.get_Range("V18");
                                PC7.Value2 = pb.FixBy;
                                Excel.Range PC8 = worksheet.get_Range("AF18");
                                PC8.Value2 = pb.CheckBy;



                            }
                            //find Count //
                            var co = db.tb_QCCountPDs.Where(c => c.WONo.Equals(WO)).ToList();

                            foreach (var rd in co)
                            {
                                if (rd.DayN.Equals("D"))
                                {
                                    if (rd.Seq <= 5)
                                    {
                                        Excel.Range CC1 = worksheet.get_Range("F" + (46 + rd.Seq).ToString());
                                        CC1.Value2 = rd.A1;
                                    }
                                    else
                                    {
                                        Excel.Range CC2 = worksheet.get_Range("R" + (46 + rd.Seq).ToString());
                                        CC2.Value2 = rd.A1;
                                    }
                                }
                                else
                                {
                                    if (rd.Seq <= 5)
                                    {
                                        Excel.Range CC1 = worksheet.get_Range("H" + (46 + rd.Seq).ToString());
                                        CC1.Value2 = rd.A1;
                                    }
                                    else
                                    {
                                        Excel.Range CC2 = worksheet.get_Range("T" + (46 + rd.Seq).ToString());
                                        CC2.Value2 = rd.A1;
                                    }
                                }
                            }

                        }
                        catch { }




                    }

                    ////////////////////////////////////////



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
        }
        public static void Print026B01032023(string WO, string PartNo, string QCNo1, string FileName)
        {
            try
            {

                string DATA = AppDomain.CurrentDomain.BaseDirectory;
                string tempPath = System.IO.Path.GetTempPath();
                string dwg = getDrawing(PartNo);
                //string FileName = "FM-PD-026.xlsx";
                ////FM-PD-026_17Feb23.xlsx
                //FileName = "FM-PD-026.07.2021.xlsx";
                //using (DataClasses1DataContext db = new DataClasses1DataContext())
                //{
                //    tb_ProductionHD pd1 = db.tb_ProductionHDs.Where(p => p.OrderNo.Equals(WO)).FirstOrDefault();
                //    if (pd1 != null)
                //    {
                //        DateTime Date1 = Convert.ToDateTime("2023-02-17 00:00:00.000");
                //        DateTime Date2 = Convert.ToDateTime("2023-03-01 00:00:00.000");

                //        if (pd1.Createdate >= Date1)
                //        {
                //            FileName = "FM-PD-026_17Feb23.xlsx";
                //        }
                //        if (pd1.Createdate >= Date2)
                //        {
                //            FileName = "FM-PD-026.01Mar23.xlsx";
                //        }

                //    }
                //}



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
                //int row1 = 22;
                //int row2 = 22;
                //int Seq = 0;
                //int seq2 = 22;
                //int CountRow = 0;
                string cIssueBy1 = "";
                string cIssueBy2 = "";
                string cIssueBy3 = "";
                string cIssueBy4 = "";

                string cCheckBy1 = "";
                string cCheckBy2 = "";
                string cCheckBy3 = "";

                string cCheckByF1 = "";
                string cCheckByF2 = "";
                string cCheckByF3 = "";

                string PV = "P";
                string QHNo = QCNo1;
                string FormISO = "";
                string DN = "";
                string SymBo = "～";
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    //string Value1 = "";
                    //string Value2 = "";
                    //string LotNo = "";



                    ///////////////SETValue/////////////////
                    var DValue = db.sp_46_QCSelectWO_01(WO).FirstOrDefault();
                    if (DValue != null)
                    {
                        DN = DValue.DayNight;
                        Excel.Range CPart = worksheet.get_Range("P3");
                        CPart.Value2 = DValue.CODE;

                        Excel.Range CStamp = worksheet.get_Range("X5");
                        if (PartNo.Length > 0)
                        {
                            if (dbClss.Right(PartNo, 1).Equals("W"))
                            {
                                CStamp.Value2 = "'" + dbClss.Right(PartNo, 8).Substring(0, 2) + "  " + dbClss.Right(PartNo, 6).Substring(0, 5);
                            }
                            else
                            {
                                CStamp.Value2 = "'" + dbClss.Right(PartNo, 7).Substring(0, 2) + "  " + dbClss.Right(PartNo, 5);
                            }
                        }


                        Excel.Range CName = worksheet.get_Range("I5");
                        CName.Value2 = DValue.NAME;

                        Excel.Range WOs = worksheet.get_Range("D5");
                        WOs.Value2 = DValue.PORDER;

                        Excel.Range CDate = worksheet.get_Range("D7");
                        CDate.Value2 = DValue.DeliveryDate;

                        Excel.Range CLot = worksheet.get_Range("D9");
                        CLot.Value2 = DValue.LotNo;

                        Excel.Range CQty = worksheet.get_Range("D11");
                        CQty.Value2 = DValue.OrderQty.ToString();


                        try
                        {
                            tb_QCHD qh = db.tb_QCHDs.Where(w => w.QCNo.Equals(QCNo1)).FirstOrDefault();
                            if (qh != null)
                            {
                                //////////Find UserName////////////
                                var uc = db.tb_QCCheckUsers.Where(u => u.QCNo.Equals(QCNo1)).ToList();
                                int r1 = 0;
                                int r2 = 0;
                                int r3 = 0;
                                int rr1 = 0;
                                int rr2 = 0;
                                int rr3 = 0;

                                foreach (var rd in uc)
                                {
                                    DN = rd.DayN;// dbShowData.CheckDayN(Convert.ToDateTime(rd.ScanDate));



                                    if (DN.Equals("D"))
                                    {
                                        if (rd.UDesc.Equals("ผู้จัดทำเอกสาร"))
                                            cIssueBy1 = rd.UserName;
                                        if (rd.UDesc.Equals("ผู้ตรวจสอบก่อนผลิต"))
                                            cIssueBy2 = rd.UserName;

                                        if (rd.UDesc.Equals("พนักงานประกอบ SUB LINE"))
                                        {
                                            r1 += 1;
                                            if (r1 > 1)
                                                cCheckBy1 += ",";
                                            cCheckBy1 += rd.UserName;
                                        }
                                        else if (rd.UDesc.Equals("พนักงานประกอบ MAIN LINE"))
                                        {
                                            r2 += 1;
                                            if (r2 > 1)
                                                cCheckBy2 += ",";
                                            cCheckBy2 += rd.UserName;
                                        }
                                        else if (rd.UDesc.Equals("พนักงานประกอบ FINAL LINE"))
                                        {
                                            r3 += 1;
                                            if (r3 > 1)
                                                cCheckBy3 += ",";
                                            cCheckBy3 += rd.UserName;
                                        }
                                    }
                                    else //N
                                    {
                                        if (rd.UDesc.Equals("ผู้จัดทำเอกสาร"))
                                            cIssueBy3 = rd.UserName;
                                        if (rd.UDesc.Equals("ผู้ตรวจสอบก่อนผลิต"))
                                            cIssueBy4 = rd.UserName;

                                        if (rd.UDesc.Equals("พนักงานประกอบ SUB LINE"))
                                        {
                                            rr1 += 1;
                                            if (rr1 > 1)
                                                cCheckByF1 += ",";
                                            cCheckByF1 += rd.UserName;


                                        }
                                        else if (rd.UDesc.Equals("พนักงานประกอบ MAIN LINE"))
                                        {
                                            rr2 += 1;
                                            if (rr2 > 1)
                                                cCheckByF2 += ",";
                                            cCheckByF2 += rd.UserName;

                                        }
                                        else if (rd.UDesc.Equals("พนักงานประกอบ FINAL LINE"))
                                        {
                                            rr3 += 1;
                                            if (rr3 > 1)
                                                cCheckByF3 += ",";
                                            cCheckByF3 += rd.UserName;

                                        }
                                    }
                                }


                                FormISO = qh.FormISO;
                                QHNo = qh.QCNo;
                                Excel.Range Ap = worksheet.get_Range("AE10");
                                Ap.Value2 = db.QC_GetUserName(qh.ApproveBy);// Convert.ToString(qh.ApproveBy);


                                Excel.Range CheckBy1 = worksheet.get_Range("E23");
                                CheckBy1.Value2 = cCheckBy1;
                                Excel.Range CheckBy2 = worksheet.get_Range("E34");
                                CheckBy2.Value2 = cCheckBy2;
                                Excel.Range CheckBy3 = worksheet.get_Range("E42");
                                CheckBy3.Value2 = cCheckBy3;

                                Excel.Range CheckByF1 = worksheet.get_Range("F23");
                                CheckByF1.Value2 = cCheckByF1;
                                Excel.Range CheckByF2 = worksheet.get_Range("F34");
                                CheckByF2.Value2 = cCheckByF2;
                                Excel.Range CheckByF3 = worksheet.get_Range("F42");
                                CheckByF3.Value2 = cCheckByF3;


                                //if (DN.Equals("D"))
                                //{
                                Excel.Range IssueBy = worksheet.get_Range("AE5");
                                IssueBy.Value2 = "1. " + cIssueBy1;
                                Excel.Range IssueBy2 = worksheet.get_Range("AE7");
                                IssueBy2.Value2 = "2. " + cIssueBy2;
                                //}
                                //else
                                //{
                                Excel.Range IssueBy3 = worksheet.get_Range("AF5");
                                IssueBy3.Value2 = "1. " + cIssueBy3;
                                Excel.Range IssueBy4 = worksheet.get_Range("AF7");
                                IssueBy4.Value2 = "2. " + cIssueBy4;
                                //}

                                QHNo = qh.QCNo;

                                ////Set Topic//

                                Excel.Range AF1 = worksheet.get_Range("AF1");
                                AF1.Value2 = getDrawing(qh.PartNo);// dwg;// "'" + db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 56);

                                //Excel.Range AA40 = worksheet.get_Range("AA40");
                                //AA40.Value2 = db.get_QC_SetDataMasterTpic(qh.FormISO, qh.PartNo, 38);

                                //Excel.Range AA32 = worksheet.get_Range("AA32");
                                //AA32.Value2 = db.get_QC_SetDataMasterTpic(qh.FormISO, qh.PartNo, 30);

                                //Step 1
                                int cRow = 43;
                                string Ppart = "";
                                for (int II = 20; II <= 22; II++)
                                {
                                    cRow += 1;

                                    ////Line 1 //
                                    Ppart = db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, II);

                                    Excel.Range Line1 = worksheet.get_Range("L" + cRow.ToString());
                                    Line1.Value2 = Ppart;

                                    var rds = db.sp_46_QCGetValue2601(qh.WONo, Ppart).FirstOrDefault();
                                    if (rds != null)
                                    {

                                        Excel.Range Line2 = worksheet.get_Range("Q" + cRow.ToString());
                                        Line2.Value2 = rds.DayN;

                                        Excel.Range Line3 = worksheet.get_Range("R" + cRow.ToString());
                                        Line3.Value2 = rds.NightN;

                                        Excel.Range Line4 = worksheet.get_Range("S" + cRow.ToString());
                                        Line4.Value2 = "'" + rds.Lot;
                                    }

                                }
                                /////FIND Part Name,Part No.
                                string Pparg2 = "";
                                //var QCP = db.sp_46_QCSelectWO_13_GroupPart(qh.WONo).ToList();
                                var QCP = db.tb_ProductionRMs.Where(p => p.OrderNo.Equals(qh.WONo)).ToList();
                                cRow = 22;
                                foreach (var rx in QCP)
                                {
                                    Ppart = "";
                                    Pparg2 = "";

                                    cRow += 1;
                                    if (cRow < 44)
                                    {

                                        Ppart = rx.PartNoRM;
                                        Pparg2 = db.getItemNo(rx.PartNoRM);
                                        if (rx.PartNoRM.ToUpper().Contains("GREASE") || rx.PartNoRM.ToUpper().Contains("LOCTITE"))
                                        {                                           

                                            Pparg2 = "";
                                        }
                                        if (Pparg2 != "")
                                        {
                                            InsertToExcel(ref worksheet, "G" + cRow.ToString(), Pparg2);
                                            InsertToExcel(ref worksheet, "L" + cRow.ToString(), Ppart);
                                            Ppart = rx.PartNoRM;
                                            var rds = db.sp_46_QCGetValue2601(qh.WONo, Ppart).FirstOrDefault();
                                            if (rds != null)
                                            {
                                                InsertToExcel(ref worksheet, "Q" + cRow.ToString(), rds.DayN);
                                                InsertToExcel(ref worksheet, "R" + cRow.ToString(), rds.NightN);
                                                InsertToExcel(ref worksheet, "S" + cRow.ToString(), rds.Lot);
                                            }
                                        }
                                    }
                                }


                                //Step 2
                                int crow2 = 22;
                                cRow = 22;
                                int NewR = 0;
                                int NewR2 = 0;
                                string CheckValueSetup = "";
                                int A35 = 0;
                                int A36 = 0;
                                int A37 = 0;
                                int A38 = 0;
                                int D23 = 0;
                                int N23 = 0;
                                for (int II = 23; II <= 57; II++)
                                {
                                    cRow += 1;
                                    crow2 += 1;
                                    CheckValueSetup = "";
                                    ////Line 1 //
                                    if (II != 20)
                                    {
                                        //การเซ็ต
                                        NewR2 = II;
                                        NewR = cRow;                                      

                                        Excel.Range Line1 = worksheet.get_Range("AE" + NewR.ToString());
                                        CheckValueSetup = db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, II);
                                        Line1.Value2 = CheckValueSetup;

                                        if (II == 50 || II==51)
                                        {
                                            Excel.Range Line2 = worksheet.get_Range("AE" + NewR.ToString());
                                            Line1.Value2 = db.QC_GetP26LotShift(qh.WONo, II);
                                        }

                                    }



                                    if (II != 44)
                                    {

                                        NewR2 = II;
                                        NewR = crow2;


                                        var rss = db.sp_46_QCGetValue2601_20(qh.WONo, II).FirstOrDefault();
                                        if (rss != null)
                                        {
                                            if (!rss.DayN.Equals(""))
                                            {
                                                Excel.Range Line2 = worksheet.get_Range("AG" + II.ToString());
                                                Line2.Value2 = rss.DayN;
                                                D23 = 1;
                                            }
                                            if (!rss.NightN.Equals(""))
                                            {
                                                Excel.Range Line3 = worksheet.get_Range("AH" + II.ToString());
                                                Line3.Value2 = rss.NightN;
                                                N23 = 1;
                                            }
                                            //Skip1//////////////////////////////////////
                                            if (CheckValueSetup.Contains("ไม่มี"))
                                            {
                                                if (D23 > 0)
                                                {
                                                    Excel.Range LineVS1 = worksheet.get_Range("AG" + II.ToString());
                                                    LineVS1.Value2 = "P";
                                                }
                                                if (N23 > 0)
                                                {
                                                    Excel.Range LineVS2 = worksheet.get_Range("AH" + II.ToString());
                                                    LineVS2.Value2 = "P";
                                                }
                                            }
                                            //Skip2//////////////////////////////////////
                                            /*
                                            if (II == 35)
                                            {
                                                if (rss.DayN.Equals("P"))
                                                {
                                                    A35 = 1;
                                                }
                                                if (rss.NightN.Equals("P"))
                                                {
                                                    A36 = 1;
                                                }
                                            }
                                            if (II == 38)
                                            {
                                                if (rss.DayN.Equals("P"))
                                                {
                                                    A37 = 1;
                                                }
                                                if (rss.NightN.Equals("P"))
                                                {
                                                    A38 = 1;
                                                }
                                            }
                                            if (II == 36)
                                            {
                                                if (A35 == 1)
                                                {
                                                    Excel.Range Line36 = worksheet.get_Range("AG" + II.ToString());
                                                    Line36.Value2 = "P";
                                                }
                                                if (A36 == 1)
                                                {
                                                    Excel.Range Line36 = worksheet.get_Range("AH" + II.ToString());
                                                    Line36.Value2 = "P";
                                                }
                                            }
                                            if (II == 39)
                                            {
                                                if (A37 == 1)
                                                {
                                                    Excel.Range Line39 = worksheet.get_Range("AG" + II.ToString());
                                                    Line39.Value2 = "P";
                                                }
                                                if (A38 == 1)
                                                {
                                                    Excel.Range Line39 = worksheet.get_Range("AH" + II.ToString());
                                                    Line39.Value2 = "P";
                                                }
                                            }
                                            */

                                            /////////////////////////////////////////////

                                        }
                                    }
                                    else
                                    {
                                        Excel.Range Line2 = worksheet.get_Range("AG44");
                                        Line2.Value2 = db.get_QC_DATAPoint_AG(qh.WONo, 44);

                                        Excel.Range Line3 = worksheet.get_Range("AE50");
                                        Line3.Value2 = db.get_QC_DATAPoint_AG(qh.WONo, 48);

                                        Excel.Range Line4 = worksheet.get_Range("AE51");
                                        Line4.Value2 = db.get_QC_DATAPoint_AG(qh.WONo, 49);
                                    }



                                }

                                Excel.Range Loctite1 = worksheet.get_Range("S44");
                                Loctite1.Value2 = "'" + db.get_QC_ValueRM(WO, "GREASE G-30M", 20);
                                Excel.Range Loctite2 = worksheet.get_Range("S45");
                                Loctite2.Value2 = "'" + db.get_QC_ValueRM(WO, "LOCTITE 277", 21);
                                Excel.Range Loctite3 = worksheet.get_Range("S46");
                                Loctite3.Value2 = "'" + db.get_QC_ValueRM(WO, "LOCTITE 414", 22);

                                string DDN1 = db.get_QC_ValueRM22(WO, "GREASE G-30M", 20);
                                string DDN2 = db.get_QC_ValueRM22(WO, "LOCTITE 277", 21);
                                string DDN3 = db.get_QC_ValueRM22(WO, "LOCTITE 414", 22);

                                //Step 1
                                if (DDN1.Equals("D") || DDN1.Equals("A"))
                                {
                                    Excel.Range LoctiteQ1 = worksheet.get_Range("Q44");
                                    LoctiteQ1.Value2 = "P";
                                }
                                if (DDN1.Equals("N") || DDN1.Equals("A"))
                                {
                                    Excel.Range LoctiteR1 = worksheet.get_Range("R44");
                                    LoctiteR1.Value2 = "P";
                                }

                                //Step 2
                                if (DDN2.Equals("D") || DDN2.Equals("A"))
                                {
                                    Excel.Range LoctiteQ2 = worksheet.get_Range("Q45");
                                    LoctiteQ2.Value2 = "P";
                                }
                                if (DDN2.Equals("N") || DDN2.Equals("A"))
                                {
                                    Excel.Range LoctiteR2 = worksheet.get_Range("R45");
                                    LoctiteR2.Value2 = "P";
                                }

                                //Step 3
                                if (DDN3.Equals("D") || DDN3.Equals("A"))
                                {
                                    Excel.Range LoctiteQ3 = worksheet.get_Range("Q46");
                                    LoctiteQ3.Value2 = "P";
                                }
                                if (DDN3.Equals("N") || DDN3.Equals("A"))
                                {
                                    Excel.Range LoctiteR3 = worksheet.get_Range("R46");
                                    LoctiteR3.Value2 = "P";
                                }

                            }
                            var gTime = db.sp_46_QCGetValue2601_Time(WO).ToList();
                            if (gTime.Count > 0)
                            {
                                var g = gTime.FirstOrDefault();
                                DateTime Chtime = Convert.ToDateTime(g.BomTime);
                                DateTime Chtime2 = Convert.ToDateTime(g.PrintTime);
                                if (g.BomTime == g.PrintTime)
                                {
                                    Chtime2 = Convert.ToDateTime(g.PrintTime).AddMinutes(30);
                                }

                                Excel.Range AB = worksheet.get_Range("AB9");
                                AB.Value2 = Math.Abs(Convert.ToDecimal((Chtime - Chtime2).TotalMinutes)).ToString("####") + " นาที";

                                if (!g.StartTime.Equals(""))
                                {
                                    Excel.Range StartT = worksheet.get_Range("N7");
                                    StartT.Value2 = Convert.ToDateTime(Chtime2).ToString("HH:mm");

                                    Excel.Range EndT = worksheet.get_Range("AA7");
                                    EndT.Value2 = Convert.ToDateTime(g.EndTime).ToString("HH:mm");

                                    // int ChanP = 0;
                                    //int.TryParse(Convert.ToInt32(DValue.ChangeModel).ToString(), out ChanP);
                                    // if (ChanP > 0)
                                    // {

                                    Excel.Range O9 = worksheet.get_Range("O9");
                                    O9.Value2 = "'" + Convert.ToDateTime(g.BomTime).ToString("HH:mm") + "-" + Convert.ToDateTime(Chtime2).ToString("HH:mm");

                                    //}

                                }
                            }

                            //Find Problem//

                            tb_QCProblem pb = db.tb_QCProblems.Where(p => p.QCNo.Equals(QHNo)).FirstOrDefault();
                            if (pb != null)
                            {
                                if (pb.TypeProblem.Equals("Man"))
                                {
                                    Excel.Range PBA = worksheet.get_Range("F13");
                                    PBA.Value2 = "P";

                                }
                                else if (pb.TypeProblem.Equals("Machine"))
                                {
                                    Excel.Range PBA = worksheet.get_Range("I13");
                                    PBA.Value2 = "P";
                                }
                                else if (pb.TypeProblem.Equals("Method"))
                                {
                                    Excel.Range PBA = worksheet.get_Range("M13");
                                    PBA.Value2 = "P";
                                }
                                else if (pb.TypeProblem.Equals("Material"))
                                {
                                    Excel.Range PBA = worksheet.get_Range("P13");
                                    PBA.Value2 = "P";
                                }
                                else if (pb.TypeProblem.Equals("Other"))
                                {
                                    Excel.Range PBA = worksheet.get_Range("S13");
                                    PBA.Value2 = "P";
                                    Excel.Range PBA2 = worksheet.get_Range("X13");
                                    PBA2.Value2 = pb.TypeRemark;
                                }

                                Excel.Range PC1 = worksheet.get_Range("F14");
                                PC1.Value2 = pb.ProblemSeeBy;
                                Excel.Range PC2 = worksheet.get_Range("N14");
                                PC2.Value2 = pb.ProblemName;

                                Excel.Range PC3 = worksheet.get_Range("AC14");
                                PC3.Value2 = pb.ProblemWare;
                                Excel.Range PC4 = worksheet.get_Range("F15");
                                PC4.Value2 = pb.ProblemTime;

                                Excel.Range PC5 = worksheet.get_Range("N15");
                                PC5.Value2 = pb.ProblemWhy;

                                Excel.Range PC6 = worksheet.get_Range("G17");
                                PC6.Value2 = pb.ProblemFix;
                                Excel.Range PC7 = worksheet.get_Range("V18");
                                PC7.Value2 = pb.FixBy;
                                Excel.Range PC8 = worksheet.get_Range("AF18");
                                PC8.Value2 = pb.CheckBy;



                            }
                            //find Count //
                            var co = db.tb_QCCountPDs.Where(c => c.WONo.Equals(WO)).ToList();

                            foreach (var rd in co)
                            {
                                if (rd.DayN.Equals("D"))
                                {
                                    if (rd.Seq <= 5)
                                    {
                                        Excel.Range CC1 = worksheet.get_Range("F" + (48 + rd.Seq).ToString());
                                        CC1.Value2 = rd.A1;
                                    }
                                    else
                                    {
                                        Excel.Range CC2 = worksheet.get_Range("R" + (48 + (rd.Seq-5)).ToString());
                                        CC2.Value2 = rd.A1;
                                    }
                                }
                                else
                                {
                                    if (rd.Seq <= 5)
                                    {
                                        Excel.Range CC1 = worksheet.get_Range("H" + (48 + (rd.Seq)).ToString());
                                        CC1.Value2 = rd.A1;
                                    }
                                    else
                                    {
                                        Excel.Range CC2 = worksheet.get_Range("T" + (48 + (rd.Seq-5)).ToString());
                                        CC2.Value2 = rd.A1;
                                    }
                                }
                            }

                        }
                        catch { }




                    }

                    ////////////////////////////////////////



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
        }

        public static void PrintData(string WO, string PartNo,string QCNo1)
        {
            string FileName = "FM-PD-026.xlsx";
            //FM-PD-026_17Feb23.xlsx
            FileName = "FM-PD-026.07.2021.xlsx";
            using (DataClasses1DataContext db = new DataClasses1DataContext())
            {
                tb_ProductionHD pd1 = db.tb_ProductionHDs.Where(p => p.OrderNo.Equals(WO)).FirstOrDefault();
                if (pd1 != null)
                {
                    DateTime Date1 = Convert.ToDateTime("2023-02-17 00:00:00.000");
                    DateTime Date2 = Convert.ToDateTime("2023-03-01 00:00:00.000");

                    if (pd1.Createdate >= Date2)
                    {

                        FileName = "FM-PD-026.01Mar23.xlsx";
                        Print026B01032023(WO, PartNo, QCNo1, FileName);
                    }
                    else if (pd1.Createdate >= Date1)
                    {
                        FileName = "FM-PD-026_17Feb23.xlsx";
                        Print026A01012020(WO, PartNo, QCNo1, FileName);
                    }
                    else
                    {
                        FileName = "FM-PD-026.07.2021.xlsx";
                        Print026A01012020(WO, PartNo, QCNo1, FileName);
                    }
                    
                }
            }

        }
        public static void PrintFMPD001(string WO, string PartNo, string QCNo1,string FromIS)
        {
            try
            {
                try
                {                   
                    string DATA = AppDomain.CurrentDomain.BaseDirectory;
                    string tempPath = System.IO.Path.GetTempPath();
                    string FileName = "FM-PD-001_STD.xlsx";

                    string TypeReport = GetReportName("STD.Base", PartNo, FromIS).ToUpper();

                    if (TypeReport.Equals("STD.PPC"))
                    {
                        FileName = "FM-PD-001_STDPPC.xlsx";
                    }
                    if (TypeReport.Equals("SPG"))
                    {
                        FileName = "FM-PD-001_SPG.xlsx";
                    }

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
                    //int row1 = 22;
                    //int row2 = 22;
                    //int Seq = 0;
                    //int seq2 = 22;
                    //int CountRow = 0;
                    string cIssueBy1 = "";
                    string cIssueBy2 = "";
                    string cIssueBy3 = "";
                    string cIssueBy4 = "";

                    string cCheckBy1 = "";
                    string cCheckBy2 = "";
                    string cCheckBy3 = "";

                    string cCheckByF1 = "";
                    string cCheckByF2 = "";
                    string cCheckByF3 = "";

                    string PV = "P";
                    string QHNo = QCNo1;
                    string FormISO = FromIS;
                    string DN = "";
                    string SymBo = "～";
                    using (DataClasses1DataContext db = new DataClasses1DataContext())
                    {
                        //string Value1 = "";
                        //string Value2 = "";
                        //string LotNo = "";

                        ///////////////SETValue/////////////////
                        var DValue = db.sp_46_QCSelectWO_01(WO).FirstOrDefault();
                        if (DValue != null)
                        {
                            DN = DValue.DayNight;

                            InsertToExcel(ref worksheet, "I5", GetSetDataDWG(PartNo, FromIS));
                            InsertToExcel(ref worksheet, "AF1", GetSetDataCust(PartNo, FromIS));
                            InsertToExcel(ref worksheet, "L48", GetSetDataLB(PartNo, FromIS, "LABEL"));
                            InsertToExcel(ref worksheet, "P3", DValue.CODE.ToString());
                            InsertToExcel(ref worksheet, "P5", DValue.NAME.ToString());
                            InsertToExcel(ref worksheet, "D5", DValue.PORDER.ToString());
                            InsertToExcel(ref worksheet, "D7", DValue.DeliveryDate);
                            InsertToExcel(ref worksheet, "D9", DValue.LotNo.ToString());
                            InsertToExcel(ref worksheet, "D11", DValue.OrderQty.ToString());
                            var gTime = db.sp_46_QCGetValue2601_Time(WO).ToList();
                            if (gTime.Count > 0)
                            {
                                var g = gTime.FirstOrDefault();
                                DateTime Chtime = Convert.ToDateTime(g.BomTime);
                                DateTime Chtime2 = Convert.ToDateTime(g.PrintTime);
                               
                                if(g.BomTime2!="")
                                    Chtime= Convert.ToDateTime(g.BomTime2);


                                if (g.BomTime == g.PrintTime)
                                {
                                    Chtime2 = Convert.ToDateTime(g.PrintTime).AddMinutes(30);
                                }
                                InsertToExcel(ref worksheet, "AB9", Math.Abs(Convert.ToDecimal((Chtime - Chtime2).TotalMinutes)).ToString("####") + " นาที");
                                if (!g.StartTime.Equals(""))
                                {
                                    InsertToExcel(ref worksheet, "N7", Convert.ToDateTime(Chtime2).ToString("HH:mm"));
                                    InsertToExcel(ref worksheet, "AA7", Convert.ToDateTime(g.EndTime).ToString("HH:mm"));
                                    InsertToExcel(ref worksheet, "O9", "'" + Convert.ToDateTime(Chtime).ToString("HH:mm") + " " + SymBo + " " + Convert.ToDateTime(Chtime2).ToString("HH:mm"));
                                }
                            }
                            try
                            {
                                tb_QCHD qh = db.tb_QCHDs.Where(w => w.QCNo.Equals(QCNo1)).FirstOrDefault();
                                if (qh != null)
                                {
                                    //////////Find UserName////////////
                                    var uc = db.tb_QCCheckUsers.Where(u => u.QCNo.Equals(QCNo1)).ToList();
                                    int r1 = 0;
                                    int r2 = 0;
                                    int r3 = 0;
                                    int rr1 = 0;
                                    int rr2 = 0;
                                    int rr3 = 0;
                                    foreach (var rd in uc)
                                    {
                                        DN = rd.DayN;// dbShowData.CheckDayN(Convert.ToDateTime(rd.ScanDate));
                                        if (DN.Equals("D"))
                                        {
                                            if (rd.UDesc.Equals("ผู้จัดทำเอกสาร"))
                                                cIssueBy1 = rd.UserName;
                                            if (rd.UDesc.Equals("ผู้ตรวจสอบก่อนผลิต"))
                                                cIssueBy2 = rd.UserName;
                                        }
                                        else //N
                                        {
                                            if (rd.UDesc.Equals("ผู้จัดทำเอกสาร"))
                                                cIssueBy3 = rd.UserName;
                                            if (rd.UDesc.Equals("ผู้ตรวจสอบก่อนผลิต"))
                                                cIssueBy4 = rd.UserName;
                                        }
                                    }

                                    InsertToExcel(ref worksheet, "AE10", db.QC_GetUserName(qh.ApproveBy));
                                    InsertToExcel(ref worksheet, "AE5", "1. " + cIssueBy1);
                                    InsertToExcel(ref worksheet, "AE7", "2. " + cIssueBy2);
                                    InsertToExcel(ref worksheet, "AF5", "1. " + cIssueBy3);
                                    InsertToExcel(ref worksheet, "AF7", "2. " + cIssueBy4);

                                    if (TypeReport.Equals("STD.PPC"))
                                    {
                                        //Main Line
                                        InsertToExcel(ref worksheet, "E23", db.get_QC_UserNameScan(qh.QCNo, "Stamp Lot /ประกอบ", "D"));
                                        InsertToExcel(ref worksheet, "F23", db.get_QC_UserNameScan(qh.QCNo, "Stamp Lot /ประกอบ", "N"));

                                        InsertToExcel(ref worksheet, "E30", db.get_QC_UserNameScan(qh.QCNo, "Test Leak", "D"));
                                        InsertToExcel(ref worksheet, "F30", db.get_QC_UserNameScan(qh.QCNo, "Test Leak", "N"));

                                        InsertToExcel(ref worksheet, "E38", db.get_QC_UserNameScan(qh.QCNo, "ตรวจสอบท้ายไลน์", "D"));
                                        InsertToExcel(ref worksheet, "F38", db.get_QC_UserNameScan(qh.QCNo, "ตรวจสอบท้ายไลน์", "N"));

                                        //Sub Line
                                        InsertToExcel(ref worksheet, "L55", db.get_QC_UserNameScan(qh.QCNo, "พนักงานตรวจสอบ SUB LINE 1", "D"));
                                        InsertToExcel(ref worksheet, "S55", db.get_QC_UserNameScan(qh.QCNo, "พนักงานตรวจสอบ SUB LINE 1", "N"));

                                        InsertToExcel(ref worksheet, "L56", db.get_QC_UserNameScan(qh.QCNo, "พนักงานตรวจสอบ SUB LINE 2", "D"));
                                        InsertToExcel(ref worksheet, "S56", db.get_QC_UserNameScan(qh.QCNo, "พนักงานตรวจสอบ SUB LINE 2", "N"));

                                        InsertToExcel(ref worksheet, "L57", db.get_QC_UserNameScan(qh.QCNo, "พนักงานประกอบ SUB LINE", "D"));
                                        InsertToExcel(ref worksheet, "S57", db.get_QC_UserNameScan(qh.QCNo, "พนักงานประกอบ SUB LINE", "N"));

                                        InsertToExcel(ref worksheet, "L58", db.get_QC_UserNameScanTime(qh.QCNo, "พนักงานประกอบ SUB LINE", "D"));
                                        InsertToExcel(ref worksheet, "S58", db.get_QC_UserNameScanTime(qh.QCNo, "พนักงานประกอบ SUB LINE", "N"));

                                        InsertToExcel(ref worksheet, "L59", db.get_QC_UserNameScan(qh.QCNo, "พนักงานขัน Plug", "D"));
                                        InsertToExcel(ref worksheet, "S59", db.get_QC_UserNameScan(qh.QCNo, "พนักงานขัน Plug", "N"));

                                        InsertToExcel(ref worksheet, "L60", db.get_QC_UserNameScanTime(qh.QCNo, "พนักงานขัน Plug", "D"));
                                        InsertToExcel(ref worksheet, "S60", db.get_QC_UserNameScanTime(qh.QCNo, "พนักงานขัน Plug", "N"));

                                        //ตอก Lot//
                                        InsertToExcel(ref worksheet, "AE33", db.get_QC_DATAPoint_AG(qh.WONo, 35));
                                        InsertToExcel(ref worksheet, "AE34", db.get_QC_DATAPoint_AG(qh.WONo, 36));


                                    }
                                    else if (TypeReport.Equals("SPG"))
                                    {
                                        //InsertToExcel(ref worksheet, "E23", db.get_QC_UserNameScan(qh.QCNo, "ขัน Plug", "D"));
                                        //InsertToExcel(ref worksheet, "F23", db.get_QC_UserNameScan(qh.QCNo, "ขัน Plug", "N"));
                                        //InsertToExcel(ref worksheet, "E27", db.get_QC_UserNameScan(qh.QCNo, "Sub Line", "D"));
                                        //InsertToExcel(ref worksheet, "F27", db.get_QC_UserNameScan(qh.QCNo, "Sub Line", "N"));
                                        //InsertToExcel(ref worksheet, "E31", db.get_QC_UserNameScan(qh.QCNo, "Stamp Lot /ประกอบ", "D"));
                                        //InsertToExcel(ref worksheet, "F31", db.get_QC_UserNameScan(qh.QCNo, "Stamp Lot /ประกอบ", "N"));

                                        //InsertToExcel(ref worksheet, "E35", db.get_QC_UserNameScan(qh.QCNo, "Test Leak", "D"));
                                        //InsertToExcel(ref worksheet, "F35", db.get_QC_UserNameScan(qh.QCNo, "Test Leak", "N"));

                                        //InsertToExcel(ref worksheet, "E39", db.get_QC_UserNameScan(qh.QCNo, "ตรวจสอบท้ายไลน์", "D"));
                                        //InsertToExcel(ref worksheet, "F39", db.get_QC_UserNameScan(qh.QCNo, "ตรวจสอบท้ายไลน์", "N"));

                                        InsertToExcel(ref worksheet, "E23", db.get_QC_UserNameScan(qh.QCNo, "Stamp Lot /ประกอบ", "D"));
                                        InsertToExcel(ref worksheet, "F23", db.get_QC_UserNameScan(qh.QCNo, "Stamp Lot /ประกอบ", "N"));

                                        InsertToExcel(ref worksheet, "E30", db.get_QC_UserNameScan(qh.QCNo, "Test Leak", "D"));
                                        InsertToExcel(ref worksheet, "F30", db.get_QC_UserNameScan(qh.QCNo, "Test Leak", "N"));

                                        InsertToExcel(ref worksheet, "E38", db.get_QC_UserNameScan(qh.QCNo, "ตรวจสอบท้ายไลน์", "D"));
                                        InsertToExcel(ref worksheet, "F38", db.get_QC_UserNameScan(qh.QCNo, "ตรวจสอบท้ายไลน์", "N"));


                                        //Sub Line
                                        InsertToExcel(ref worksheet, "L55", db.get_QC_UserNameScan(qh.QCNo, "พนักงานตรวจสอบ SUB LINE 1", "D"));
                                        InsertToExcel(ref worksheet, "S55", db.get_QC_UserNameScan(qh.QCNo, "พนักงานตรวจสอบ SUB LINE 1", "N"));

                                        InsertToExcel(ref worksheet, "L56", db.get_QC_UserNameScan(qh.QCNo, "พนักงานตรวจสอบ SUB LINE 2", "D"));
                                        InsertToExcel(ref worksheet, "S56", db.get_QC_UserNameScan(qh.QCNo, "พนักงานตรวจสอบ SUB LINE 2", "N"));

                                        InsertToExcel(ref worksheet, "L57", db.get_QC_UserNameScan(qh.QCNo, "พนักงานประกอบ SUB LINE", "D"));
                                        InsertToExcel(ref worksheet, "S57", db.get_QC_UserNameScan(qh.QCNo, "พนักงานประกอบ SUB LINE", "N"));
                                        InsertToExcel(ref worksheet, "L58", db.get_QC_UserNameScanTime(qh.QCNo, "พนักงานประกอบ SUB LINE", "D"));
                                        InsertToExcel(ref worksheet, "S58", db.get_QC_UserNameScanTime(qh.QCNo, "พนักงานประกอบ SUB LINE", "N"));

                                        InsertToExcel(ref worksheet, "L59", db.get_QC_UserNameScan(qh.QCNo, "พนักงานขัน Plug", "D"));
                                        InsertToExcel(ref worksheet, "S59", db.get_QC_UserNameScan(qh.QCNo, "พนักงานขัน Plug", "N"));
                                        InsertToExcel(ref worksheet, "L60", db.get_QC_UserNameScanTime(qh.QCNo, "พนักงานขัน Plug", "D"));
                                        InsertToExcel(ref worksheet, "S60", db.get_QC_UserNameScanTime(qh.QCNo, "พนักงานขัน Plug", "N"));

                                        //ตอก Lot//
                                        InsertToExcel(ref worksheet, "AE33", db.get_QC_DATAPoint_AG(qh.WONo, 35));
                                        InsertToExcel(ref worksheet, "AE34", db.get_QC_DATAPoint_AG(qh.WONo, 36));
                                        

                                    }
                                    else if (TypeReport.ToUpper().Equals("STD.BASE"))
                                    {
                                        InsertToExcel(ref worksheet, "E23", db.get_QC_UserNameScan(qh.QCNo, "Stamp Lot /ประกอบ", "D"));
                                        InsertToExcel(ref worksheet, "F23", db.get_QC_UserNameScan(qh.QCNo, "Stamp Lot /ประกอบ", "N"));

                                        InsertToExcel(ref worksheet, "E30", db.get_QC_UserNameScan(qh.QCNo, "Test Leak", "D"));
                                        InsertToExcel(ref worksheet, "F30", db.get_QC_UserNameScan(qh.QCNo, "Test Leak", "N"));

                                        InsertToExcel(ref worksheet, "E38", db.get_QC_UserNameScan(qh.QCNo, "ตรวจสอบท้ายไลน์", "D"));
                                        InsertToExcel(ref worksheet, "F38", db.get_QC_UserNameScan(qh.QCNo, "ตรวจสอบท้ายไลน์", "N"));

                                        //Sub Line
                                        InsertToExcel(ref worksheet, "L55", db.get_QC_UserNameScan(qh.QCNo, "พนักงานตรวจสอบ SUB LINE 1", "D"));
                                        InsertToExcel(ref worksheet, "S55", db.get_QC_UserNameScan(qh.QCNo, "พนักงานตรวจสอบ SUB LINE 1", "N"));

                                        InsertToExcel(ref worksheet, "L56", db.get_QC_UserNameScan(qh.QCNo, "พนักงานตรวจสอบ SUB LINE 2", "D"));
                                        InsertToExcel(ref worksheet, "S56", db.get_QC_UserNameScan(qh.QCNo, "พนักงานตรวจสอบ SUB LINE 2", "N"));

                                        InsertToExcel(ref worksheet, "L57", db.get_QC_UserNameScan(qh.QCNo, "พนักงานประกอบ SUB LINE", "D"));
                                        InsertToExcel(ref worksheet, "S57", db.get_QC_UserNameScan(qh.QCNo, "พนักงานประกอบ SUB LINE", "N"));
                                        InsertToExcel(ref worksheet, "L58", db.get_QC_UserNameScanTime(qh.QCNo, "พนักงานประกอบ SUB LINE", "D"));
                                        InsertToExcel(ref worksheet, "S58", db.get_QC_UserNameScanTime(qh.QCNo, "พนักงานประกอบ SUB LINE", "N"));

                                        InsertToExcel(ref worksheet, "L59", db.get_QC_UserNameScan(qh.QCNo, "พนักงานขัน Plug", "D"));
                                        InsertToExcel(ref worksheet, "S59", db.get_QC_UserNameScan(qh.QCNo, "พนักงานขัน Plug", "N"));

                                        InsertToExcel(ref worksheet, "L60", db.get_QC_UserNameScanTime(qh.QCNo, "พนักงานขัน Plug", "D"));
                                        InsertToExcel(ref worksheet, "S60", db.get_QC_UserNameScanTime(qh.QCNo, "พนักงานขัน Plug", "N"));

                                        //ตอก Lot//
                                        InsertToExcel(ref worksheet, "AE33", db.get_QC_DATAPoint_AG(qh.WONo, 35));
                                        InsertToExcel(ref worksheet, "AE34", db.get_QC_DATAPoint_AG(qh.WONo, 36));
                                      

                                    }
                                    //Find CountPD
                                    /////////////////////////////

                                    /////////////////////////////
                                    var co = db.tb_QCCountPDs.Where(c => c.WONo.Equals(WO)).ToList();
                                    foreach (var rd in co)
                                    {
                                        if (rd.DayN.Equals("D"))
                                        {
                                            InsertToExcel(ref worksheet, "F" + (62 + rd.Seq).ToString(), rd.A1);
                                        }
                                        else
                                        {
                                            InsertToExcel(ref worksheet, "H" + (62 + rd.Seq).ToString(), rd.A1);
                                        }

                                    }
                                    //Find Problem//

                                    tb_QCProblem pb = db.tb_QCProblems.Where(p => p.QCNo.Equals(QHNo)).FirstOrDefault();
                                    if (pb != null)
                                    {
                                        if (pb.TypeProblem.Equals("Man"))
                                        {
                                            InsertToExcel(ref worksheet, "F13", "P");
                                        }
                                        else if (pb.TypeProblem.Equals("Machine"))
                                        {
                                            InsertToExcel(ref worksheet, "F13", "P");
                                        }
                                        else if (pb.TypeProblem.Equals("Method"))
                                        {
                                            InsertToExcel(ref worksheet, "M13", "P");
                                        }
                                        else if (pb.TypeProblem.Equals("Material"))
                                        {
                                            InsertToExcel(ref worksheet, "P13", "P");
                                        }
                                        else if (pb.TypeProblem.Equals("Other"))
                                        {
                                            InsertToExcel(ref worksheet, "S13", "P");
                                            InsertToExcel(ref worksheet, "X13", pb.TypeRemark);
                                        }
                                        InsertToExcel(ref worksheet, "E14", pb.ProblemSeeBy);
                                        InsertToExcel(ref worksheet, "N14", pb.ProblemName);
                                        InsertToExcel(ref worksheet, "AC14", pb.ProblemWare);
                                        InsertToExcel(ref worksheet, "E15", pb.ProblemTime);
                                        InsertToExcel(ref worksheet, "N15", pb.ProblemWhy);
                                        InsertToExcel(ref worksheet, "E17", pb.ProblemFix);
                                        InsertToExcel(ref worksheet, "V18", pb.FixBy);
                                        InsertToExcel(ref worksheet, "AE18", pb.CheckBy);

                                    }
                                    ////Scan Time///
                                    var urList = db.tb_QCCheckUserTimes.Where(p => p.QCNo.Equals(QCNo1) && !p.BoxNo.Equals("")).ToList();
                                    if (urList.Count > 0)
                                    {
                                        foreach (var rd in urList)
                                        {

                                            if (rd.UDesc.Equals("ประกอบ"))
                                            {
                                                InsertToExcel(ref worksheet, CheckColumnTime(rd.BoxNo) + "51", rd.UserName);
                                            }
                                            else if (rd.UDesc.Equals("Test Leak"))
                                            {
                                                InsertToExcel(ref worksheet, CheckColumnTime(rd.BoxNo) + "52", rd.UserName);
                                            }
                                            else
                                            {
                                                InsertToExcel(ref worksheet, CheckColumnTime(rd.BoxNo) + "53", rd.UserName);
                                            }
                                        }
                                    }
                                    ////////////////

                                    ////Set Topic//
                                    //Step 1
                                    int cRow = 22;
                                    int IR = 0;
                                    string Ppart = "";
                                    string Pparg2 = "";

                                    //for (int II = 1; II <= 26; II++)
                                    //{

                                    //    cRow += 1;
                                    //    if (II <26)
                                    //    {
                                    //        Ppart = db.get_QC_SetDataMasterX(qh.FormISO, qh.PartNo, II - 2);
                                    //        Pparg2 = db.get_QC_SetDataMasterX(qh.FormISO, qh.PartNo, II - 2);

                                    //        InsertToExcel(ref worksheet, "G" + cRow.ToString(), Ppart);
                                    //        InsertToExcel(ref worksheet, "L" + cRow.ToString(), Pparg2);

                                    //        var rds = db.sp_46_QCGetValue2601(qh.WONo, Ppart).FirstOrDefault();
                                    //        if (rds != null)
                                    //        {
                                    //            InsertToExcel(ref worksheet, "Q" + cRow.ToString(), rds.DayN);
                                    //            InsertToExcel(ref worksheet, "R" + cRow.ToString(), rds.NightN);
                                    //            InsertToExcel(ref worksheet, "S" + cRow.ToString(), rds.Lot);
                                    //        }
                                    //    }                                    

                                    //}
                                    //  var QCP = db.sp_46_QCSelectWO_13_GroupPart(qh.WONo).ToList();
                                    var QCP1 = db.tb_ProductionRMs.Where(p => p.OrderNo.Equals(qh.WONo)).ToList();
                                    cRow = 22;
                                    foreach (var rx in QCP1)
                                    {
                                        Ppart = "";
                                        Pparg2 = "";

                                        cRow += 1;
                                        if (cRow < 48)
                                        {
                                            // MOLYBDENUM GREASE (S - GREASE)                                           
                                            Ppart = rx.PartNoRM;
                                            Pparg2 = db.getItemNo(rx.PartNoRM);
                                            if (rx.PartNoRM.ToUpper().Contains("MOLYBDENUM") || rx.PartNoRM.ToUpper().Contains("DYNAMAX") || rx.PartNoRM.ToUpper().Contains("LOCTITE"))
                                            {
                                                // Pparg2 = rx.PartNo;
                                                //if (rx.PartNo.ToUpper().Contains("MOLYBDENUM"))                                                    
                                                //     Ppart = db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 21);
                                                //if (rx.PartNo.ToUpper().Contains("DYNAMAX"))
                                                //    Ppart = db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 22);
                                                //if (rx.PartNo.ToUpper().Contains("LOCTITE"))
                                                //    Ppart = db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 23);

                                                Pparg2 = "";
                                            }
                                            if (Pparg2 != "")
                                            {
                                                InsertToExcel(ref worksheet, "G" + cRow.ToString(), Pparg2);
                                                InsertToExcel(ref worksheet, "L" + cRow.ToString(), Ppart);
                                                Ppart = rx.PartNoRM;
                                                var rds = db.sp_46_QCGetValue2601(qh.WONo, Ppart).FirstOrDefault();
                                                if (rds != null)
                                                {
                                                    InsertToExcel(ref worksheet, "Q" + cRow.ToString(), rds.DayN);
                                                    InsertToExcel(ref worksheet, "R" + cRow.ToString(), rds.NightN);
                                                    InsertToExcel(ref worksheet, "S" + cRow.ToString(), rds.Lot);
                                                }
                                            }
                                        }
                                    }
                                    Ppart = "";
                                    Ppart = db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 24);
                                    if (!Ppart.Equals(""))
                                    {
                                        InsertToExcel(ref worksheet, "L48", Ppart);
                                    }

                                    string Ppart1 = "";
                                    string Ppart2 = "";
                                    string Ppart3 = "";
                                    Ppart1 = db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 21);
                                    Ppart2 = db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 22);
                                    Ppart3 = db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 23);
                                    InsertToExcel(ref worksheet, "L45", Ppart1);
                                    InsertToExcel(ref worksheet, "L46", Ppart2);
                                    InsertToExcel(ref worksheet, "L47", Ppart3);
                                    Ppart1 = "";
                                    Ppart2 = "";
                                    Ppart3 = "";
                                    Ppart1 = db.get_QC_SetDataMaster2(qh.FormISO, qh.PartNo, 21);
                                    Ppart2 = db.get_QC_SetDataMaster2(qh.FormISO, qh.PartNo, 22);
                                    Ppart3 = db.get_QC_SetDataMaster2(qh.FormISO, qh.PartNo, 23);
                                    InsertToExcel(ref worksheet, "G45", Ppart1);
                                    InsertToExcel(ref worksheet, "G46", Ppart2);
                                    InsertToExcel(ref worksheet, "G47", Ppart3);

                                    var rds1 = db.sp_46_QCGetValue2601(qh.WONo, Ppart1).FirstOrDefault();
                                    if (rds1 != null)
                                    {
                                        InsertToExcel(ref worksheet, "Q45", rds1.DayN);
                                        InsertToExcel(ref worksheet, "R45", rds1.NightN);
                                        InsertToExcel(ref worksheet, "S45", rds1.Lot);
                                    }

                                    var rds2 = db.sp_46_QCGetValue2601(qh.WONo, Ppart2).FirstOrDefault();
                                    if (rds2 != null)
                                    {
                                        InsertToExcel(ref worksheet, "Q46", rds2.DayN);
                                        InsertToExcel(ref worksheet, "R46", rds2.NightN);
                                        InsertToExcel(ref worksheet, "S46", rds2.Lot);
                                    }

                                    var rds3 = db.sp_46_QCGetValue2601(qh.WONo, Ppart3).FirstOrDefault();
                                    if (rds3 != null)
                                    {
                                        InsertToExcel(ref worksheet, "Q47", rds3.DayN);
                                        InsertToExcel(ref worksheet, "R47", rds3.NightN);
                                        InsertToExcel(ref worksheet, "S47", rds3.Lot);
                                    }

                                    /////Step 2
                                    ////  int crow2 = 22;
                                    cRow = 22;
                                    int CK = 0;
                                    int N23 = 0;
                                    int D23 = 0;
                                    for (int II = 25; II <= 74; II++)
                                    {

                                        cRow += 1;
                                        // crow2 += 1;


                                        if (TypeReport.Equals("STD.PPC"))
                                        {
                                            if (II == 44)
                                            {
                                                //  II += 3;
                                                cRow += 2;
                                            }

                                            Ppart = "";
                                            Ppart = db.get_QC_SetDataMasterP01(qh.FormISO, qh.PartNo, II);
                                            if (!Ppart.ToLower().Equals("xnonex"))
                                            {
                                                if (("AE" + cRow.ToString()).Equals("AE39") || ("AE" + cRow.ToString()).Equals("AE33") || ("AE" + cRow.ToString()).Equals("AE34"))
                                                {

                                                }
                                                else
                                                {

                                                    InsertToExcel(ref worksheet, "AE" + cRow.ToString(), Ppart);
                                                }
                                            }
                                            if (II == 37)
                                                cRow = 35;
                                            if (II == 38)
                                                cRow = 37;
                                            if (II == 39)
                                                cRow = 39;
                                        }
                                        else
                                        {

                                            if (II == 43)
                                                cRow = 43;
                                            Ppart = "";
                                            Ppart = db.get_QC_SetDataMasterP01(qh.FormISO, qh.PartNo, II);
                                            if (!Ppart.ToLower().Equals("xnonex"))
                                            {

                                                if (("AE" + cRow.ToString()).Equals("AE45") || ("AE" + cRow.ToString()).Equals("AE33") || ("AE" + cRow.ToString()).Equals("AE34"))
                                                {

                                                }
                                                else
                                                {
                                                    InsertToExcel(ref worksheet, "AE" + cRow.ToString(), Ppart);
                                                }

                                            }


                                        }

                                        ////Line 1 //
                                        CK = Convert.ToInt32(db.get_QC_DATAPoint_AG_2D(qh.WONo, qh.PartNo, qh.FormISO, II));
                                        if (CK == 1)
                                        {
                                            InsertToExcel(ref worksheet, "AG" + cRow.ToString(), "P");
                                            D23 += 1;
                                        }
                                        else if (CK == 2)
                                        {
                                            InsertToExcel(ref worksheet, "AH" + cRow.ToString(), "P");
                                            N23 += 1;
                                        }
                                        else if (CK == 3)
                                        {
                                            if (D23 > 0)
                                            {
                                                InsertToExcel(ref worksheet, "AG" + cRow.ToString(), "P");
                                            }
                                            if (N23 > 23)
                                            {
                                                InsertToExcel(ref worksheet, "AH" + cRow.ToString(), "P");
                                            }
                                        }
                                        else
                                        {

                                        }
                                    }

                                }
                                string LLOT = qh.LotNo;
                                if (TypeReport.Equals("STD.PPC"))
                                {
                                    LLOT = db.get_QC_DATAPoint_AG(qh.WONo, 39);
                                    if (LLOT=="") { LLOT = qh.LotNo; }
                                    InsertToExcel(ref worksheet, "AE39", "LOT ( " + LLOT + " )");
                                    InsertToExcel(ref worksheet, "AE46", db.get_QC_DATAPoint_AG(qh.WONo, 44));
                                    InsertToExcel(ref worksheet, "AE47", db.get_QC_DATAPoint_AG(qh.WONo, 45));
                                }
                                else if (TypeReport.Equals("SPG"))
                                {
                                    LLOT = db.get_QC_DATAPoint_AG(qh.WONo, 45);
                                    if (LLOT == "") { LLOT = qh.LotNo; }
                                    InsertToExcel(ref worksheet, "AE45", "LOT ( " + LLOT + " )");
                                    InsertToExcel(ref worksheet, "AE46", db.get_QC_DATAPoint_AG(qh.WONo, 46));
                                    InsertToExcel(ref worksheet, "AE47", db.get_QC_DATAPoint_AG(qh.WONo, 47));
                                }
                                else if (TypeReport.Equals("STD.Base"))
                                {
                                    LLOT = db.get_QC_DATAPoint_AG(qh.WONo, 45);
                                    if (LLOT == "") { LLOT = qh.LotNo; }
                                    InsertToExcel(ref worksheet, "AE45", "LOT ( " + LLOT + " )");
                                    InsertToExcel(ref worksheet, "AE46", db.get_QC_DATAPoint_AG(qh.WONo, 46));
                                    InsertToExcel(ref worksheet, "AE47", db.get_QC_DATAPoint_AG(qh.WONo, 47));
                                }
                            
                            }
                            catch { }
                        }
                        ////////////////////////////////////////
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
            }
            catch { }
        }
        public static void PrintFMPD002(string WO, string PartNo, string QCNo1, string FromIS)
        {
            try
            {                
                string DATA = AppDomain.CurrentDomain.BaseDirectory;
                string tempPath = System.IO.Path.GetTempPath();
                string FileName = "FM-PD-002_STD.xlsx";
                string TypeReport = GetReportName("STD", PartNo, FromIS).ToUpper();
                string HeaderText = "100% Inspection Check Sheet BRAKE CHAMBER";
                if (TypeReport.Equals("SPG"))
                {
                    HeaderText = "100% Inspection Check Sheet SPRING CHAMBER";
                }
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
                Excel.Worksheet worksheet2 = (Excel.Worksheet)sheets.get_Item(2);
                Excel.Worksheet worksheet3 = (Excel.Worksheet)sheets.get_Item(3);
                Excel.Worksheet worksheet4 = (Excel.Worksheet)sheets.get_Item(4);
                // progressBar1.Maximum = 51;
                // progressBar1.Minimum = 1;
                int row1 = 6;
                int Seq = 0;
                int TG = 0;
                string PV = "P";
                string QHNo = QCNo1;
                string FormISO = "";
                string [] Inspector= new string[10];
                string cCheckBy1 = "";
                string cCheckBy2 = "";
                string cCheckBy3 = "";
                string cCheckBy4 = "";
                string cCheckBy5 = "";
                string cCheckBy6 = "";
                string[] SetData = new string[10];
                bool PAGE1 = true;
                bool PAGE2 = false;
                bool PAGE3 = false;
                bool PAGE4 = false;
                bool chek24 = true;
                string DN = "";
                string LotMark = "ตอกที่ Base มี4 ตำแหน่ง , มีการตอก Lot ถูกต้อง , ชัดเจน , สามารถอ่านได้ ";
                string Line1Part = "";
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    string RefValue1 = "";
                    string RefValue2 = "";
                    string RefValue3 = "";
                    string PartName = "";
                    string C9 = "";
                    string GP5 = "";
                    ///////////////SETValue/////////////////////////////////
                    var DValue = db.sp_46_QCSelectWO_01(WO).FirstOrDefault();
                    if (DValue != null)
                    {
                        LotMark = LotMark + " ( "+DValue.LotNo+" )";
                        var PTAGList = db.tb_QCTAGs.Where(p => p.QCNo.Equals(QHNo)).ToList();
                        if (PTAGList.Count > 40)
                        {
                            PAGE2 = true;
                        }
                        if (PTAGList.Count > 80)
                        {
                            PAGE3 = true;
                        }
                        if (PTAGList.Count > 120)
                        {
                            PAGE4 = true;
                        }
                        DN = DValue.DayNight;
                        PartName = DValue.NAME;

                       

                        if (PAGE1)
                        {
                            InsertToExcel(ref worksheet, "A3", HeaderText);
                            InsertToExcel(ref worksheet, "S3", DValue.CODE);
                            InsertToExcel(ref worksheet, "S4", DValue.NAME);
                            InsertToExcel(ref worksheet, "Q5", DValue.PORDER);
                            InsertToExcel(ref worksheet, "Y5", DValue.LotNo);                            
                        }
                        if (PAGE2)
                        {
                            InsertToExcel(ref worksheet2, "A3", HeaderText);
                            InsertToExcel(ref worksheet2, "S3", DValue.CODE);
                            InsertToExcel(ref worksheet2, "S4", DValue.NAME);
                            InsertToExcel(ref worksheet2, "Q5", DValue.PORDER);
                            InsertToExcel(ref worksheet2, "Y5", DValue.LotNo);
                        }
                        if (PAGE3)
                        {
                            InsertToExcel(ref worksheet3, "A3", HeaderText);
                            InsertToExcel(ref worksheet3, "S3", DValue.CODE);
                            InsertToExcel(ref worksheet3, "S4", DValue.NAME);
                            InsertToExcel(ref worksheet3, "Q5", DValue.PORDER);
                            InsertToExcel(ref worksheet3, "Y5", DValue.LotNo);
                        }
                        if (PAGE4)
                        {
                            InsertToExcel(ref worksheet4, "A3", HeaderText);
                            InsertToExcel(ref worksheet4, "S3", DValue.CODE);
                            InsertToExcel(ref worksheet4, "S4", DValue.NAME);
                            InsertToExcel(ref worksheet4, "Q5", DValue.PORDER);
                            InsertToExcel(ref worksheet4, "Y5", DValue.LotNo);
                        }          
                        

                        try
                        {
                            tb_QCHD qh = db.tb_QCHDs.Where(w => w.QCNo.Equals(QCNo1)).FirstOrDefault();
                            if (qh != null)
                            {
                                //////////Find UserName////////////
                                var uc = db.tb_QCCheckUsers.Where(u => u.QCNo.Equals(QCNo1)).ToList();
                                foreach (var rd in uc)
                                {
                                    DN = rd.DayN;
                                    if (DN.Equals("D"))
                                    {
                                        if (rd.UDesc.Equals("Check By"))
                                        {
                                            if (cCheckBy1.Equals(""))
                                                cCheckBy1 = rd.UserName;
                                            else
                                                cCheckBy1 = cCheckBy1 + "/" + rd.UserName;                                            
                                        }
                                       
                                    }
                                    else
                                    {
                                        if (rd.UDesc.Equals("Check By"))
                                        {

                                            if (cCheckBy4.Equals(""))
                                                cCheckBy4 = rd.UserName;
                                            else
                                                cCheckBy4 = cCheckBy4 + "/" + rd.UserName;
                                            // DN1 = dbShowData.CheckDayN(Convert.ToDateTime(rd.ScanDate));
                                        }
                                    
                                    }
                                }
                                FormISO = qh.FormISO;
                                QHNo = qh.QCNo;
                                RefValue1 = qh.RefValue1;
                                RefValue2 = qh.RefValue2;
                                RefValue3 = qh.RefValue3;
                                string E5 = "";
                                int RowS = 6;
                                var GroupPartList = db.tb_QCGroupParts.Where(p => p.PartNo.Equals(PartNo) && p.FormISO.Equals(FromIS) && p.GroupPart.Trim().Equals("")).OrderBy(s=>s.Seq).ToList();

                                if (PAGE1)
                                {
                                    InsertToExcel(ref worksheet, "AD4", db.QC_GetUserName(qh.ApproveBy));
                                    InsertToExcel(ref worksheet, "AN4", cCheckBy1);
                                    InsertToExcel(ref worksheet, "AQ4", cCheckBy4);
                                    //Step ตรวจสอบ//
                                    InsertToExcel(ref worksheet, "Y25", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 1", 0));
                                    InsertToExcel(ref worksheet, "AC25", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 2", 0));
                                    InsertToExcel(ref worksheet, "AG25", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 3", 0));
                                    InsertToExcel(ref worksheet, "AK25", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 4", 0));
                                    InsertToExcel(ref worksheet, "AO25", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 5", 0));

                                    InsertToExcel(ref worksheet, "Y26", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 1", 1));
                                    InsertToExcel(ref worksheet, "AC26", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 2", 1));
                                    InsertToExcel(ref worksheet, "AG26", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 3", 1));
                                    InsertToExcel(ref worksheet, "AK26", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 4", 1));
                                    InsertToExcel(ref worksheet, "AO26", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 5", 1));
                                    //End Step//
                                    E5 = Convert.ToDateTime(qh.CreateDate).ToString("dd") + " วัน " + Convert.ToDateTime(qh.CreateDate).ToString("MM") + " เดือน  " + Convert.ToDateTime(qh.CreateDate).ToString("yyyy") + " ปี";
                                    InsertToExcel(ref worksheet, "E5", E5);
                                    InsertToExcel(ref worksheet, "D4", qh.OrderQty); //Prod. Qty
                                    InsertToExcel(ref worksheet, "E4", db.get_QCSumQtyNG_RE(qh.QCNo, 1)); // NG
                                    InsertToExcel(ref worksheet, "G4", db.get_QCSumQtyNG_RE(qh.QCNo, 2)); //Rework
                                    
                                    
                                      
                                    RowS = 6;
                                    foreach(var rd in GroupPartList)
                                    {
                                        RowS += 1;
                                        InsertToExcel(ref worksheet, "A"+(RowS).ToString(), rd.Seq);
                                        InsertToExcel(ref worksheet, "B" + (RowS).ToString(), rd.TopPic);                                      
                                        Line1Part = rd.SetData;
                                        if (rd.TopPic.ToUpper().Equals("LOT"))
                                        {
                                            Line1Part = LotMark;
                                        }
                                        InsertToExcel(ref worksheet, "C" + (RowS).ToString(), Line1Part);                                        
                                    }                                 
                                }

                                if (PAGE2)
                                {
                                    InsertToExcel(ref worksheet2, "AD4", db.QC_GetUserName(qh.ApproveBy));
                                    InsertToExcel(ref worksheet2, "AN4", cCheckBy1);
                                    InsertToExcel(ref worksheet2, "AQ4", cCheckBy4);
                                    //Step ตรวจสอบ//
                                    InsertToExcel(ref worksheet2, "Y25", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 1", 0));
                                    InsertToExcel(ref worksheet2, "AC25", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 2", 0));
                                    InsertToExcel(ref worksheet2, "AG25", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 3", 0));
                                    InsertToExcel(ref worksheet2, "AK25", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 4", 0));
                                    InsertToExcel(ref worksheet2, "AO25", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 5", 0));

                                    InsertToExcel(ref worksheet2, "Y26", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 1", 1));
                                    InsertToExcel(ref worksheet2, "AC26", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 2", 1));
                                    InsertToExcel(ref worksheet2, "AG26", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 3", 1));
                                    InsertToExcel(ref worksheet2, "AK26", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 4", 1));
                                    InsertToExcel(ref worksheet2, "AO26", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 5", 1));
                                    //End Step//
                                    E5 = Convert.ToDateTime(qh.CreateDate).ToString("dd") + " วัน " + Convert.ToDateTime(qh.CreateDate).ToString("MM") + " เดือน  " + Convert.ToDateTime(qh.CreateDate).ToString("yyyy") + " ปี";
                                    InsertToExcel(ref worksheet2, "E5", E5);
                                    InsertToExcel(ref worksheet2, "D4", qh.OrderQty); //Prod. Qty
                                    InsertToExcel(ref worksheet2, "E4", db.get_QCSumQtyNG_RE(qh.QCNo, 1)); // NG
                                    InsertToExcel(ref worksheet2, "G4", db.get_QCSumQtyNG_RE(qh.QCNo, 2)); //Rework
                                    //InsertToExcel(ref worksheet, "I4", E5); //OK
                                    //InsertToExcel(ref worksheet, "K4", E5); //Net NG    
                                    RowS = 6;
                                    foreach (var rd in GroupPartList)
                                    {
                                        RowS += 1;
                                        InsertToExcel(ref worksheet2, "A" + (RowS).ToString(), rd.Seq);
                                        InsertToExcel(ref worksheet2, "B" + (RowS).ToString(), rd.TopPic);
                                        Line1Part = rd.SetData;
                                        if (rd.TopPic.ToUpper().Equals("LOT"))
                                        {
                                            Line1Part = LotMark;
                                        }
                                        InsertToExcel(ref worksheet2, "C" + (RowS).ToString(), Line1Part);
                                    }
                                }
                                if (PAGE3)
                                {
                                    InsertToExcel(ref worksheet3, "AD4", db.QC_GetUserName(qh.ApproveBy));
                                    InsertToExcel(ref worksheet3, "AN4", cCheckBy1);
                                    InsertToExcel(ref worksheet3, "AQ4", cCheckBy4);
                                    //Step ตรวจสอบ//
                                    InsertToExcel(ref worksheet3, "Y25", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 1", 0));
                                    InsertToExcel(ref worksheet3, "AC25", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 2", 0));
                                    InsertToExcel(ref worksheet3, "AG25", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 3", 0));
                                    InsertToExcel(ref worksheet3, "AK25", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 4", 0));
                                    InsertToExcel(ref worksheet3, "AO25", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 5", 0));

                                    InsertToExcel(ref worksheet3, "Y26", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 1", 1));
                                    InsertToExcel(ref worksheet3, "AC26", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 2", 1));
                                    InsertToExcel(ref worksheet3, "AG26", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 3", 1));
                                    InsertToExcel(ref worksheet3, "AK26", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 4", 1));
                                    InsertToExcel(ref worksheet3, "AO26", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 5", 1));
                                    //End Step//
                                    E5 = Convert.ToDateTime(qh.CreateDate).ToString("dd") + " วัน " + Convert.ToDateTime(qh.CreateDate).ToString("MM") + " เดือน  " + Convert.ToDateTime(qh.CreateDate).ToString("yyyy") + " ปี";
                                    InsertToExcel(ref worksheet3, "E5", E5);
                                    InsertToExcel(ref worksheet3, "D4", qh.OrderQty); //Prod. Qty
                                    InsertToExcel(ref worksheet3, "E4", db.get_QCSumQtyNG_RE(qh.QCNo, 1)); // NG
                                    InsertToExcel(ref worksheet3, "G4", db.get_QCSumQtyNG_RE(qh.QCNo, 2)); //Rework
                                    //InsertToExcel(ref worksheet, "I4", E5); //OK
                                    //InsertToExcel(ref worksheet, "K4", E5); //Net NG    
                                    RowS = 6;
                                    foreach (var rd in GroupPartList)
                                    {
                                        RowS += 1;
                                        InsertToExcel(ref worksheet3, "A" + (RowS).ToString(), rd.Seq);
                                        InsertToExcel(ref worksheet3, "B" + (RowS).ToString(), rd.TopPic);
                                        Line1Part = rd.SetData;
                                        if (rd.TopPic.ToUpper().Equals("LOT"))
                                        {
                                            Line1Part = LotMark;
                                        }
                                        InsertToExcel(ref worksheet3, "C" + (RowS).ToString(), Line1Part);
                                    }
                                }
                                if (PAGE4)
                                {
                                    InsertToExcel(ref worksheet4, "AD4", db.QC_GetUserName(qh.ApproveBy));
                                    InsertToExcel(ref worksheet4, "AN4", cCheckBy1);
                                    InsertToExcel(ref worksheet4, "AQ4", cCheckBy4);
                                    //Step ตรวจสอบ//
                                    InsertToExcel(ref worksheet4, "Y25", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 1", 0));
                                    InsertToExcel(ref worksheet4, "AC25", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 2", 0));
                                    InsertToExcel(ref worksheet4, "AG25", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 3", 0));
                                    InsertToExcel(ref worksheet4, "AK25", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 4", 0));
                                    InsertToExcel(ref worksheet4, "AO25", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 5", 0));

                                    InsertToExcel(ref worksheet4, "Y26", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 1", 1));
                                    InsertToExcel(ref worksheet4, "AC26", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 2", 1));
                                    InsertToExcel(ref worksheet4, "AG26", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 3", 1));
                                    InsertToExcel(ref worksheet4, "AK26", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 4", 1));
                                    InsertToExcel(ref worksheet4, "AO26", db.get_QC_CheckUserBoxNo(qh.QCNo, "ผู้ตรวจสอบ 5", 1));
                                    //End Step//

                                    E5 = Convert.ToDateTime(qh.CreateDate).ToString("dd") + " วัน " + Convert.ToDateTime(qh.CreateDate).ToString("MM") + " เดือน  " + Convert.ToDateTime(qh.CreateDate).ToString("yyyy") + " ปี";
                                    InsertToExcel(ref worksheet4, "E5", E5);
                                    InsertToExcel(ref worksheet4, "D4", qh.OrderQty); //Prod. Qty
                                    InsertToExcel(ref worksheet4, "E4", db.get_QCSumQtyNG_RE(qh.QCNo, 1)); // NG
                                    InsertToExcel(ref worksheet4, "G4", db.get_QCSumQtyNG_RE(qh.QCNo, 2)); //Rework
                                    //InsertToExcel(ref worksheet, "I4", E5); //OK
                                    //InsertToExcel(ref worksheet, "K4", E5); //Net NG    
                                    RowS = 6;
                                    foreach (var rd in GroupPartList)
                                    {
                                        RowS += 1;
                                        InsertToExcel(ref worksheet4, "A" + (RowS).ToString(), rd.Seq);
                                        InsertToExcel(ref worksheet4, "B" + (RowS).ToString(), rd.TopPic);
                                        Line1Part = rd.SetData;
                                        if (rd.TopPic.ToUpper().Equals("LOT"))
                                        {
                                            Line1Part = LotMark;
                                        }
                                        InsertToExcel(ref worksheet4, "C" + (RowS).ToString(), Line1Part);
                                    }
                                }
                            }

                        }
                        catch (Exception ex) { MessageBox.Show("first " + ex.Message); }
                    }

                    ////////////////////////////////////////

                    int countA = 0;
                    int CountB = 0;
                    int CountC = 0;
                    int CountD = 0;
                    int TAG2 = 0;
                    int CA = 0;
                    int TG2 = 0;
                    int NGA = 0;
                    int NGB = 0;
                    int NGC = 0;
                    string TAGOf1 = "";
                    string TAGOf2 = "";
                    string TAGOf3 = "";

                    int CountTAG = 0;
                    var listPoint = db.sp_46_QCSelectWO_09_QCTAGSelect(QHNo).ToList();                                       
                    CountTAG = listPoint.Count;
                    if (listPoint.Count > 0)
                    {
                        foreach (var rs in listPoint)
                        {
                            countA += 1;
                            if (countA > 40)
                            {
                                CountB += 1;
                            }
                            if (countA > 80)
                            {
                                CountC += 1;
                            }
                            if (countA > 120)
                            {
                                CountD += 1;
                            }
                            TG = 0;

                            string[] PPTAG = rs.BarcodeTag.Split(',');
                            TG = Convert.ToInt32(PPTAG[2]);
                            TAG2 += TG;                           
                            TG2 = 0;
                            TG2 = Convert.ToInt32(db.get_QCSumQtyTAGNG(QHNo, rs.BarcodeTag, 3));

                            if (listPoint.Count == countA)
                            {
                                NGA = TG;
                                TAGOf1 = PPTAG[5];
                            }
                            if ((listPoint.Count - 1) == countA)
                            {
                                NGB = TG;
                                TAGOf2 = PPTAG[5];
                            }
                            if ((listPoint.Count - 2) == countA)
                            {
                                NGC = TG;
                                TAGOf3 = PPTAG[5];
                            }
                            row1 = 6;
                            Seq = 0;
                            var listPart = db.tb_QCGroupParts.Where(q => q.FormISO.Equals(FromIS) && q.PartNo.Equals(DValue.CODE) && q.GroupPart.Trim().Equals("")).OrderBy(o => o.Seq).ToList();
                            CA = listPart.Count();
                            foreach (var rd in listPart)
                            {
                                row1 += 1;
                                Seq += 1;
                                if (!rd.SetData.Equals("") && row1 <= 18)
                                {
                                    try
                                    {
                                        var gValue = db.sp_46_QCGetValue5601(rs.BarcodeTag, QHNo, rd.Seq).FirstOrDefault();
                                        PV = "P";
                                        if (gValue.CountA > 0)
                                        {
                                            PV = "O";

                                            if (gValue.CountA == 99)
                                            {
                                                PV = "";
                                            }
                                        }
                                        //if (rd.Seq.Equals(9) && C9.Equals(""))
                                        //{
                                        //    PV = "";
                                        //}
                                        if (countA <= 40)
                                        {                                          
                                            InsertToExcel(ref worksheet, Getcolumn(countA + 4) + row1.ToString(), PV);
                                        }
                                        if (countA > 40 && countA <= 80)
                                        {
                                            InsertToExcel(ref worksheet2, Getcolumn(CountB + 4) + row1.ToString(), PV);
                                        }
                                        if (countA > 80 && countA <= 120)
                                        {                                           
                                            InsertToExcel(ref worksheet3, Getcolumn(CountC + 4) + row1.ToString(), PV);
                                        }
                                        if (countA > 120 && countA <= 160)
                                        {
                                            InsertToExcel(ref worksheet4, Getcolumn(CountD + 4) + row1.ToString(), PV);
                                        }
                                    }
                                    catch { }
                                    //catch (Exception ex) { MessageBox.Show("Mid " + ex.Message); }

                                }
                                //SumNG//       

                                if (countA <= 40)
                                {                 
                                    InsertToExcel(ref worksheet, Getcolumn(countA + 4) + "20", (TG - TG2));
                                    InsertToExcel(ref worksheet, Getcolumn(countA + 4) + "21", TG2);
                                }
                                else if (countA > 40 && countA <= 80)
                                {
                                    if (PAGE2)
                                    {
                                        InsertToExcel(ref worksheet2, Getcolumn(CountB + 4) + "20", (TG - TG2));
                                        InsertToExcel(ref worksheet2, Getcolumn(CountB + 4) + "21", TG2);
                                    }
                                }
                                else if (countA > 80 && countA <= 120)
                                {
                                    if (PAGE3)
                                    {
                                        InsertToExcel(ref worksheet3, Getcolumn(CountC + 4) + "20", (TG - TG2));
                                        InsertToExcel(ref worksheet3, Getcolumn(CountC + 4) + "21", TG2);

                                    }
                                }
                                else if (countA > 120 && countA <= 160)
                                {
                                    if (PAGE4)
                                    {
                                        InsertToExcel(ref worksheet4, Getcolumn(CountD + 4) + "20", (TG - TG2));
                                        InsertToExcel(ref worksheet4, Getcolumn(CountD + 4) + "21", TG2);
                                    }
                                }

                            }//foreach 
                             //}//cunt A //Page 1 End
                        }//for

                        ////NGQty and Remark//
                        int RM = 0;
                        int TNG1 = 0;
                        int TNG2 = 0;
                        int TNG3 = 0;
                        tb_QCProblem qcp = db.tb_QCProblems.Where(p => p.QCNo.Equals(QHNo) && !p.NGQty.Equals(0)).FirstOrDefault();
                        if (qcp != null)
                        {
                            var tgf = db.tb_QCTAGs.Where(s => s.QCNo.Equals(QHNo)).ToList();
                            foreach (var tf in tgf)
                            {
                                if (tf.ofTAG.Equals(TAGOf1))
                                {
                                    TNG1 = Convert.ToInt32(tf.NGQty);
                                }
                                else if (tf.ofTAG.Equals(TAGOf2))
                                {
                                    TNG2 = Convert.ToInt32(tf.NGQty);
                                }
                                else if (tf.ofTAG.Equals(TAGOf3))
                                {
                                    TNG3 = Convert.ToInt32(tf.NGQty);
                                }
                            }

                            if (countA <= 40)
                            {   //NG Qty//                
                                InsertToExcel(ref worksheet, "C19", qcp.ProblemName);
                                ///////////////////////////////
                                if (qcp.NGQty > 0)
                                {
                                    if (NGA < (qcp.NGQty + TNG1))
                                    {                                        
                                        InsertToExcel(ref worksheet, Getcolumn(countA + 4) + "19", "O");
                                        InsertToExcel(ref worksheet, Getcolumn(countA + 4) + "20", 0);
                                        InsertToExcel(ref worksheet, Getcolumn(countA + 4) + "21",NGA);
                                        RM = (Convert.ToInt32(qcp.NGQty) + TNG1) - NGA;
                                        if (RM > 0)
                                        {
                                            if (NGB < (RM + TNG2))
                                            {                                                
                                                InsertToExcel(ref worksheet, Getcolumn((countA -1) + 4) + "19", "O");
                                                InsertToExcel(ref worksheet, Getcolumn((countA -1) + 4) + "20", 0);
                                                InsertToExcel(ref worksheet, Getcolumn((countA -1) + 4) + "21", NGB);
                                                RM = (RM + TNG2) - NGB;
                                                if (RM > 0)
                                                {
                                                    if (NGC < (RM + TNG3))
                                                    {                                                        
                                                        InsertToExcel(ref worksheet, Getcolumn((countA - 2) + 4) + "19", "O");
                                                        InsertToExcel(ref worksheet, Getcolumn((countA - 2) + 4) + "20", 0);
                                                        InsertToExcel(ref worksheet, Getcolumn((countA - 2) + 4) + "21", NGC);
                                                    }
                                                    else
                                                    {                                                        
                                                        InsertToExcel(ref worksheet, Getcolumn((countA - 2) + 4) + "19", "O");
                                                        InsertToExcel(ref worksheet, Getcolumn((countA - 2) + 4) + "20", NGC - (RM + TNG3));
                                                        InsertToExcel(ref worksheet, Getcolumn((countA - 2) + 4) + "21", RM + TNG3);
                                                    }
                                                }
                                            }
                                            else
                                            {                                                
                                                InsertToExcel(ref worksheet, Getcolumn((countA - 1) + 4) + "19", "O");
                                                InsertToExcel(ref worksheet, Getcolumn((countA - 1) + 4) + "20", NGB - (RM + TNG2));
                                                InsertToExcel(ref worksheet, Getcolumn((countA - 1) + 4) + "21", (RM + TNG2));
                                            }
                                        }


                                    }
                                    else
                                    {                                     
                                        InsertToExcel(ref worksheet, Getcolumn((countA - 0) + 4) + "19", "O");
                                        InsertToExcel(ref worksheet, Getcolumn((countA - 0) + 4) + "20", NGA - (qcp.NGQty + TNG1));
                                        InsertToExcel(ref worksheet, Getcolumn((countA - 0) + 4) + "21", (qcp.NGQty + TNG1));

                                    }

                                }

                            }
                            else if (countA > 40 && countA <= 80)
                            {
                                if (PAGE2)
                                {
                                    //NG Qty//                                   
                                    InsertToExcel(ref worksheet2, "C19", qcp.ProblemName);
                                    if (qcp.NGQty > 0)
                                    {
                                        if (NGA < (qcp.NGQty + TNG1))
                                        {                                           
                                            InsertToExcel(ref worksheet2, Getcolumn(countA + 4) + "19", "O");
                                            InsertToExcel(ref worksheet2, Getcolumn(countA + 4) + "20", 0);
                                            InsertToExcel(ref worksheet2, Getcolumn(countA + 4) + "21", NGA);
                                            RM = (Convert.ToInt32(qcp.NGQty) + TNG1) - NGA;
                                            if (RM > 0)
                                            {
                                                if (NGB < (RM + TNG2))
                                                {
                                                    
                                                    InsertToExcel(ref worksheet2, Getcolumn((countA - 1) + 4) + "19", "O");
                                                    InsertToExcel(ref worksheet2, Getcolumn((countA - 1) + 4) + "20", 0);
                                                    InsertToExcel(ref worksheet2, Getcolumn((countA - 1) + 4) + "21", NGB);

                                                    RM = (RM + TNG2) - NGB;
                                                    if (RM > 0)
                                                    {
                                                        if (NGC < (RM + TNG3))
                                                        {
                                                            
                                                            InsertToExcel(ref worksheet2, Getcolumn((countA - 2) + 4) + "19", "O");
                                                            InsertToExcel(ref worksheet2, Getcolumn((countA - 2) + 4) + "20", 0);
                                                            InsertToExcel(ref worksheet2, Getcolumn((countA - 2) + 4) + "21", NGC);
                                                        }
                                                        else
                                                        {                                                            
                                                            InsertToExcel(ref worksheet2, Getcolumn((countA - 2) + 4) + "19", "O");
                                                            InsertToExcel(ref worksheet2, Getcolumn((countA - 2) + 4) + "20", NGC - (RM + TNG3));
                                                            InsertToExcel(ref worksheet2, Getcolumn((countA - 2) + 4) + "21", RM + TNG3);
                                                        }
                                                    }
                                                }
                                                else
                                                {                                                    
                                                    InsertToExcel(ref worksheet2, Getcolumn((countA - 1) + 4) + "19", "O");
                                                    InsertToExcel(ref worksheet2, Getcolumn((countA - 1) + 4) + "20", NGB - (RM + TNG2));
                                                    InsertToExcel(ref worksheet2, Getcolumn((countA - 1) + 4) + "21", (RM + TNG2));
                                                }
                                            }


                                        }
                                        else
                                        {                                           
                                            InsertToExcel(ref worksheet2, Getcolumn((countA - 0) + 4) + "19", "O");
                                            InsertToExcel(ref worksheet2, Getcolumn((countA - 0) + 4) + "20", NGA - (qcp.NGQty + TNG1));
                                            InsertToExcel(ref worksheet2, Getcolumn((countA - 0) + 4) + "21", (qcp.NGQty + TNG1));
                                        }

                                    }
                                }
                            }
                            else if (countA > 80 && countA <= 120)
                            {
                                if (PAGE3)
                                {
                                    //NG Qty//                                   
                                    InsertToExcel(ref worksheet3, "C19", qcp.ProblemName);
                                    if (qcp.NGQty > 0)
                                    {
                                        if (NGA < (qcp.NGQty + TNG1))
                                        {                                            
                                            InsertToExcel(ref worksheet3, Getcolumn(countA + 4) + "19", "O");
                                            InsertToExcel(ref worksheet3, Getcolumn(countA + 4) + "20", 0);
                                            InsertToExcel(ref worksheet3, Getcolumn(countA + 4) + "21", NGA);
                                            RM = (Convert.ToInt32(qcp.NGQty) + TNG1) - NGA;
                                            if (RM > 0)
                                            {
                                                if (NGB < (RM + TNG2))
                                                {                                                   

                                                    InsertToExcel(ref worksheet3, Getcolumn((countA - 1) + 4) + "19", "O");
                                                    InsertToExcel(ref worksheet3, Getcolumn((countA - 1) + 4) + "20", 0);
                                                    InsertToExcel(ref worksheet3, Getcolumn((countA - 1) + 4) + "21", NGB);
                                                    RM = (RM + TNG2) - NGB;
                                                    if (RM > 0)
                                                    {
                                                        if (NGC < (RM + TNG3))
                                                        {                                                           
                                                            InsertToExcel(ref worksheet3, Getcolumn((countA - 2) + 4) + "19", "O");
                                                            InsertToExcel(ref worksheet3, Getcolumn((countA - 2) + 4) + "20", 0);
                                                            InsertToExcel(ref worksheet3, Getcolumn((countA - 2) + 4) + "21", NGC);
                                                        }
                                                        else
                                                        {                                                           
                                                            InsertToExcel(ref worksheet3, Getcolumn((countA - 2) + 4) + "19", "O");
                                                            InsertToExcel(ref worksheet3, Getcolumn((countA - 2) + 4) + "20", NGC - (RM + TNG3));
                                                            InsertToExcel(ref worksheet3, Getcolumn((countA - 2) + 4) + "21", RM + TNG3);
                                                        }
                                                    }
                                                }
                                                else
                                                {                                                    
                                                    InsertToExcel(ref worksheet3, Getcolumn((countA - 1) + 4) + "19", "O");
                                                    InsertToExcel(ref worksheet3, Getcolumn((countA - 1) + 4) + "20", NGB - (RM + TNG2));
                                                    InsertToExcel(ref worksheet3, Getcolumn((countA - 1) + 4) + "21", (RM + TNG2));
                                                }
                                            }


                                        }
                                        else
                                        {                                            
                                            InsertToExcel(ref worksheet3, Getcolumn((countA - 0) + 4) + "19", "O");
                                            InsertToExcel(ref worksheet3, Getcolumn((countA - 0) + 4) + "20", NGA - (qcp.NGQty + TNG1));
                                            InsertToExcel(ref worksheet3, Getcolumn((countA - 0) + 4) + "21", (qcp.NGQty + TNG1));
                                        }

                                    }
                                }
                            }
                            else if (countA > 120 && countA <= 160)
                            {
                                if (PAGE4)
                                {
                                    //NG Qty//                                    
                                    InsertToExcel(ref worksheet4, "C19", qcp.ProblemName);

                                    if (qcp.NGQty > 0)
                                    {
                                        if (NGA < (qcp.NGQty + TNG1))
                                        {                                           
                                            InsertToExcel(ref worksheet4, Getcolumn(countA + 4) + "19", "O");
                                            InsertToExcel(ref worksheet4, Getcolumn(countA + 4) + "20", 0);
                                            InsertToExcel(ref worksheet4, Getcolumn(countA + 4) + "21", NGA);
                                            RM = (Convert.ToInt32(qcp.NGQty) + TNG1) - NGA;
                                            if (RM > 0)
                                            {
                                                if (NGB < (RM + TNG2))
                                                {                                                    
                                                    InsertToExcel(ref worksheet4, Getcolumn((countA - 1) + 4) + "19", "O");
                                                    InsertToExcel(ref worksheet4, Getcolumn((countA - 1) + 4) + "20", 0);
                                                    InsertToExcel(ref worksheet4, Getcolumn((countA - 1) + 4) + "21", NGB);
                                                    RM = (RM + TNG2) - NGB;
                                                    if (RM > 0)
                                                    {
                                                        if (NGC < (RM + TNG3))
                                                        {                                                           
                                                            InsertToExcel(ref worksheet4, Getcolumn((countA - 2) + 4) + "19", "O");
                                                            InsertToExcel(ref worksheet4, Getcolumn((countA - 2) + 4) + "20", 0);
                                                            InsertToExcel(ref worksheet4, Getcolumn((countA - 2) + 4) + "21", NGC);
                                                        }
                                                        else
                                                        {                                                            
                                                            InsertToExcel(ref worksheet4, Getcolumn((countA - 2) + 4) + "19", "O");
                                                            InsertToExcel(ref worksheet4, Getcolumn((countA - 2) + 4) + "20", NGC - (RM + TNG3));
                                                            InsertToExcel(ref worksheet4, Getcolumn((countA - 2) + 4) + "21", RM + TNG3);
                                                        }
                                                    }
                                                }
                                                else
                                                {                                                   
                                                    InsertToExcel(ref worksheet4, Getcolumn((countA - 1) + 4) + "19", "O");
                                                    InsertToExcel(ref worksheet4, Getcolumn((countA - 1) + 4) + "20", NGB - (RM + TNG2));
                                                    InsertToExcel(ref worksheet4, Getcolumn((countA - 1) + 4) + "21", (RM + TNG2));
                                                }
                                            }


                                        }
                                        else
                                        {                                            
                                            InsertToExcel(ref worksheet4, Getcolumn((countA - 0) + 4) + "19", "O");
                                            InsertToExcel(ref worksheet4, Getcolumn((countA - 0) + 4) + "20", NGA - (qcp.NGQty + TNG1));
                                            InsertToExcel(ref worksheet4, Getcolumn((countA - 0) + 4) + "21", (qcp.NGQty + TNG1));
                                        }

                                    }
                                }
                            }
                        }

                        ////// PC Check ///
                    }



                }

                excelBook.SaveAs(tempfile);
                excelBook.Close(false);
                excelApp.Quit();
                releaseObject(worksheet);
                releaseObject(worksheet2);
                releaseObject(worksheet3);
                releaseObject(excelBook);
                releaseObject(excelApp);

                Marshal.FinalReleaseComObject(worksheet);
                Marshal.FinalReleaseComObject(worksheet2);
                Marshal.FinalReleaseComObject(worksheet3);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet2);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet3);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                GC.GetTotalMemory(false);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.GetTotalMemory(true);
                System.Diagnostics.Process.Start(tempfile);

            }
            catch (Exception ex) { MessageBox.Show("last " + ex.Message); }
        }
        public static void PrintFMPD003(string WO, string PartNo, string QCNo1, string FromIS)
        {
            try
            {
                string DATA = AppDomain.CurrentDomain.BaseDirectory;
                string tempPath = System.IO.Path.GetTempPath();
                string FileName = "FM-PD-003_STD.xlsx";
                string tempfile = tempPath + FileName;
               

                string TypeReport = GetReportName("STD", PartNo, FromIS).ToUpper();
                string HeaderText = " Shipping Check  result in process  ( Brake chamber)";
                if (TypeReport.Equals("SPG"))
                {
                    HeaderText = " Shipping Check  result in process  ( SPG B/ CH-W)";
                    FileName = "FM-PD-003_SPG.xlsx";
                }
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
                int row1 = 14;

                string PV = "P";
                string QHNo = QCNo1;
                string FormISO = FromIS;
                string DN = "";                
                string cCheckBy1 = "";
                string cCheckBy2 = "";
                string cCheckBy3 = "";
                string LotRemark = "";
                string E5 = "";
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {   string PartName = "";
                    string LotNo = "";
                    ///////////////SETValue/////////////////
                    var DValue = db.sp_46_QCSelectWO_01(WO).FirstOrDefault();
                    if (DValue != null)
                    {
                        DN = DValue.DayNight;
                        PartName = DValue.NAME;
                        InsertToExcel(ref worksheet, "A4", HeaderText);
                        InsertToExcel(ref worksheet, "C6", DValue.NAME);
                        InsertToExcel(ref worksheet, "C7", DValue.CODE);
                        InsertToExcel(ref worksheet, "C8", db.get_QC_ItemCustomerNo(DValue.CODE,""));                        
                        E5 = " วัน "+ Convert.ToDateTime(DValue.CreateDate).ToString("dd") + " เดือน " + Convert.ToDateTime(DValue.CreateDate).ToString("MM") + " ปี  " + Convert.ToDateTime(DValue.CreateDate).ToString("yyyy") + "";
                        InsertToExcel(ref worksheet, "C9", E5);
                        InsertToExcel(ref worksheet, "C10",DValue.OrderQty);
                        InsertToExcel(ref worksheet, "C11", DValue.LotNo);
                        InsertToExcel(ref worksheet, "D13", DValue.PORDER);
                        LotRemark = "ตอก Lot ที่ Base ต้องไม่ลืมตอก, ไม่ตอกผิดตำแหน่งที่กำหนดไว้  มีการตอก Lot ถูกต้อง , ชัดเจน , สามารถอ่านได้ ( " + DValue.LotNo + " ) ";
                        try
                        {
                            tb_QCHD qh = db.tb_QCHDs.Where(w => w.QCNo.Equals(QCNo1)).FirstOrDefault();
                            if (qh != null)
                            {
                                FormISO = qh.FormISO;
                                LotNo = qh.LotNo;
                                //////////Find UserName////////////
                                var uc = db.tb_QCCheckUsers.Where(u => u.QCNo.Equals(QCNo1)).ToList();
                                foreach (var rd in uc)
                                {
                                    DN = dbShowData.CheckDayN(Convert.ToDateTime(rd.ScanDate));

                                    if (rd.UDesc.Equals("ผู้ตรวจสอบ 1"))
                                        cCheckBy1 = rd.UserName;
                                    if (rd.UDesc.Equals("ผู้ตรวจสอบ 2"))
                                        cCheckBy2 = rd.UserName;
                                    if (rd.UDesc.Equals("ผู้ตรวจสอบ 3"))
                                        cCheckBy3 = rd.UserName;
                                }


                                InsertToExcel(ref worksheet, "J8", cCheckBy1);
                                InsertToExcel(ref worksheet, "K8", cCheckBy2);
                                InsertToExcel(ref worksheet, "H7", db.QC_GetUserName(qh.ApproveBy));
                                QHNo = qh.QCNo;

                                if (db.QC_CheckNG(QHNo) == "P")
                                {
                                    InsertToExcel(ref worksheet, "J10", "P");
                                }
                                else
                                {
                                    InsertToExcel(ref worksheet, "H10", "P");
                                }

                                //Excel.Range O19 = worksheet.get_Range("O19");
                                //O19.Value2 = db.get_QC_DATAPoint(qh.QCNo, "", 8);
                                //Excel.Range O20 = worksheet.get_Range("O20");
                                //O20.Value2 = db.get_QC_DATAPoint(qh.QCNo, "", 9);
                            }

                        }
                        catch { }

                    }

                    ////////////////////////////////////////
                    ////Insert Desc.//
                    int countA = 0;
                    row1 = 14;
                    string setPart = "";
                    var vList = db.tb_QCGroupParts.Where(p => p.FormISO.Equals(FromIS) && p.PartNo.Equals(PartNo) && !p.TopPic.Trim().Equals("")
                                && p.GroupPart.Trim().Equals("")).OrderBy(s => s.Seq).ToList();
                    foreach(var ls in vList)
                    {
                        countA += 1;
                        row1 += 1;
                        setPart = ls.SetData.Trim();
                        if(ls.TopPic.Trim().Equals("Lot No."))
                        {
                            setPart = LotRemark;
                        }
                        InsertToExcel(ref worksheet, "A"+ row1.ToString(), countA);
                        InsertToExcel(ref worksheet, "B" + row1.ToString(), ls.TopPic.Trim());
                        InsertToExcel(ref worksheet, "D" + row1.ToString(), setPart);

                    }

                    ////////////////


                    countA = 0;
                    string Colm = "";
                    //var ListQ = db.tb_QCTAGs.Where(p => p.QCNo.Equals(QCNo1)).ToList();
                    //int SQ = 0;
                    //foreach(var re in ListQ)
                    //{
                    //    if(re.BarcodeTag.Contains("Head"))
                    //    {
                    //        SQ = 1;
                    //    }
                    //    else if(re.BarcodeTag.Equals(""))
                    //    {
                    //        SQ = 2;
                    //    }
                    //    else
                    //    {
                    //        SQ = 3;

                    //    }
                    //    tb_QCTAG pUpdate = db.tb_QCTAGs.Where(s => s.id.Equals(re.id)).FirstOrDefault();
                    //    if(pUpdate!=null)
                    //    {
                    //        pUpdate.TAGQ = SQ;
                    //        db.SubmitChanges();
                    //    }
                          
                    //}
                    var listPoint = db.sp_46_QCSelectWO_09_QCTAGSelectPD(QHNo).ToList();                    
                    if (listPoint.Count > 0)
                    {
                        foreach (var rs in listPoint)
                        {
                            countA += 1; 
                            if (countA <= 3)
                            {
                                row1 = 14;
                                var listPart = db.tb_QCGroupParts.Where(p => p.FormISO.Equals(FromIS) && p.PartNo.Equals(PartNo) && !p.TopPic.Trim().Equals("")
                                                && p.GroupPart.Trim().Equals("")).OrderBy(s => s.Seq).ToList();
                                foreach (var rd in listPart)
                                {                                   
                                    row1 += 1;                                   
                                    if (!rd.SetData.Equals("") && row1 <= 40)
                                    {
                                        try
                                        {

                                            var gValue = db.sp_46_QCGetValue5601(rs.BarcodeTag, QHNo, rd.Seq).FirstOrDefault();
                                            PV = "OK";

                                            if (gValue.CountA > 0)
                                            {
                                                PV = "OK";
                                                if (gValue.CountA > 0)
                                                    PV = "NG";
                                                if (gValue.CountA == 99)
                                                    PV = "";
                                            }



                                            if (rs.BarcodeTag.Contains("Head"))
                                            {
                                                Colm = "H";
                                            }
                                            else if (rs.BarcodeTag.Contains("Middle"))
                                            {
                                                Colm = "J";
                                            }else
                                            {
                                                Colm = "K";
                                            }
                                           
                                        

                                            if (TypeReport.ToUpper().Equals("SPG"))
                                            {
                                                if(row1>=15 && row1<=22)
                                                {
                                                    //2-8
                                                    //input digital
                                                    if (row1 != 15)
                                                        PV = db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);
                                                    if(Colm.Equals("J") || Colm.Equals("K"))
                                                    {
                                                        if (row1 >= 15 && row1 <= 19)
                                                        {
                                                            PV = "";
                                                        }
                                                    }
                                                }
                                               
                                            }
                                            else
                                            {
                                                if (row1 >= 15 && row1 <19)
                                                {
                                                    //2-4
                                                    //Input digital
                                                    if(row1!=15)
                                                         PV = db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);

                                                    if (Colm.Equals("J") || Colm.Equals("K"))
                                                    {
                                                        if (row1 >= 15 && row1 <= 18)
                                                        {
                                                            PV = "";
                                                        }
                                                    }
                                                }
                                            }
                                        //Input Value///
                                          InsertToExcel(ref worksheet, Colm + row1.ToString(), PV); 
                                        }
                                        catch (Exception ex) { MessageBox.Show(ex.Message); }
                                    }

                                }//foreach
                            }//cunt A
                        }//for
                    }



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
        }
        public static void PrintData5601(string WO, string PartNo, string QCNo1)
        {
            try
            {


                string DATA = AppDomain.CurrentDomain.BaseDirectory;
                string tempPath = System.IO.Path.GetTempPath();
                string FileName = "FM-QA-056.xlsx";
                // FileName = "FM-QA-056_02_1.xlsx";
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    
                    tb_ProductionHD pd = db.tb_ProductionHDs.Where(p => p.OrderNo.Equals(WO) 
                    && (p.LineName2.Equals("TW10-CB") || p.LineName2.Equals("TS10-SC_CB") || p.LineName2.Equals("TW02-SC_PB"))).FirstOrDefault();
                    if (pd != null)
                    {
                        FileName = "FM-QA-056_CM.xlsx";
                        PrintData5601CM(WO, PartNo, QCNo1);
                        return;
                    }
                }

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
                int row1 = 8;
                int row2 = 9;
                int Seq = 0;
                int seq2 = 21;
                int CountRow = 0;
                string PV = "P";
                string QHNo = QCNo1;
                string FormISO = "";
                string DN = "";
                string cIssueBy1 = "";
                string cCheckBy1 = "";


                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    string Value1 = "";
                    string Value2 = "";
                    string LotNo = "";
                    string RefValue1 = "";
                    string I6 = "P";
                    string L6 = "";
                    string ValueInvalid = "";
                    ///////////////SETValue/////////////////
                    var DValue = db.sp_46_QCSelectWO_01(WO).FirstOrDefault();
                    if (DValue != null)
                    {

                        //Sampling inspection report BRAKE CHAMBER
                        //Sampling inspection report piggyback

                        DN = DValue.DayNight;
                        Excel.Range CPart = worksheet.get_Range("C3");
                        CPart.Value2 = DValue.NAME;
                        Excel.Range CStamp = worksheet.get_Range("C2");
                        CStamp.Value2 = DValue.CODE;
                        Excel.Range CName = worksheet.get_Range("C4");
                        CName.Value2 = DValue.OrderQty;

                        Excel.Range CDate = worksheet.get_Range("C5");
                        CDate.Value2 = DValue.LotNo;

                        Excel.Range CWO = worksheet.get_Range("C6");
                        CWO.Value2 = WO;


                        try
                        {
                            tb_QCHD qh = db.tb_QCHDs.Where(w => w.QCNo.Equals(QCNo1)).FirstOrDefault();
                            if (qh != null)
                            {

                                Excel.Range App = worksheet.get_Range("I3");
                                App.Value2 = db.QC_GetUserName(qh.ApproveBy);
                                if (!qh.ApproveBy.Equals(""))
                                {
                                    Excel.Range Appdate = worksheet.get_Range("I5");
                                    Appdate.Value2 = qh.ApproveDate;
                                }

                                QHNo = qh.QCNo;
                                RefValue1 = qh.RefValue1;
                                FormISO = qh.FormISO;
                                //////////Find UserName////////////
                                var uc = db.tb_QCCheckUsers.Where(u => u.QCNo.Equals(QHNo)).ToList();


                                foreach (var rd in uc)
                                {
                                    DN = dbShowData.CheckDayN(Convert.ToDateTime(rd.ScanDate));

                                    if (rd.UDesc.Equals("Inspector"))
                                    {
                                        cIssueBy1 = rd.UserName;
                                        Excel.Range K3 = worksheet.get_Range("L3");
                                        K3.Value2 = cIssueBy1;
                                        Excel.Range K5 = worksheet.get_Range("L5");
                                        K5.Value2 = rd.ScanDate;
                                    }
                                    //if (rd.UDesc.Equals("Check By"))
                                    //{
                                    //    cCheckBy1 = rd.UserName;
                                    //    Excel.Range I3 = worksheet.get_Range("K3");
                                    //    I3.Value2 = cCheckBy1;
                                    //    Excel.Range I5 = worksheet.get_Range("K5");
                                    //    I5.Value2 = rd.ScanDate;

                                    //}
                                }
                                //Pass/Not Pass
                                if (!qh.ApproveBy.Equals(""))
                                {
                                    if (db.QC_CheckNG(QHNo) == "P")
                                    {
                                        Excel.Range L6x = worksheet.get_Range("L6");
                                        L6x.Value2 = L6;
                                    }
                                    else
                                    {
                                        Excel.Range I6x = worksheet.get_Range("I6");
                                        I6x.Value2 = I6;
                                    }
                                }

                            }

                        }
                        catch (Exception ex) { MessageBox.Show("1." + ex.Message); }




                    }

                    ////////////////////////////////////////
                    int countA = 0;
                    string col = "";
                    string col2x = "";
                    var listPoint = db.sp_46_QCSelectWO_09_QCTAGSelect(QHNo).ToList();
                    if (listPoint.Count > 0)
                    {
                        foreach (var rs in listPoint)
                        {
                            countA += 1;
                            // MessageBox.Show(countA.ToString());
                            if (countA <= 2)
                            {
                                row1 = 9;
                                col = "I";
                                col2x = "G";
                                if (countA == 2)
                                {
                                    col = "L";
                                    col2x = "H";
                                }



                                var listPart = db.tb_QCGroupParts.Where(q => q.FormISO.Equals(FormISO) && q.PartNo.Equals(DValue.CODE)).OrderBy(o => o.Seq).ToList();
                                foreach (var rd in listPart)
                                {
                                    //Start Insert Checkmark                            

                                    //if (rd.Seq <= 14)
                                    //{
                                    row1 += 1;

                                    Excel.Range SetDT = worksheet.get_Range("D" + row1.ToString());
                                    SetDT.Value2 = db.get_QC_SetDataMaster(FormISO, rd.PartNo, rd.Seq);
                                    //Start G=7,H=
                                    if (!rd.SetData.Equals(""))
                                    {
                                        try
                                        {
                                            var gValue = db.sp_46_QCGetValue5601(rs.BarcodeTag, QHNo, rd.Seq).FirstOrDefault();

                                            PV = "P";
                                            if (gValue.CountA > 0)
                                            {
                                                PV = "O";
                                                if (gValue.CountA == 99)
                                                    PV = "";
                                            }
                                            if (countA == 2 && row1 == 10)
                                                PV = "";

                                            if (rd.Seq >= 6 && rd.Seq <= 7)
                                            {
                                                PV = "";
                                                Excel.Range Col02 = worksheet.get_Range(col2x + row1.ToString());
                                                Col02.Value2 = db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);

                                                Excel.Range Col0K = worksheet.get_Range(col + row1.ToString());
                                                string ValueP = db.get_QC_DATAPointValue4(QHNo, rs.BarcodeTag, rd.Seq);

                                                if (ValueP == "OK")
                                                {
                                                    Col0K.Value2 = "P";
                                                }
                                                else if (ValueP == "NG")
                                                {
                                                    Col0K.Value2 = "O";
                                                    I6 = "";
                                                    L6 = "P";
                                                }

                                            }
                                            else if (rd.Seq == 8)
                                            {

                                                if (db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq) != "")
                                                {
                                                    if (ValueInvalid == "")
                                                        ValueInvalid = db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);
                                                    else
                                                        ValueInvalid = ValueInvalid + "," + db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);

                                                    Excel.Range Col8 = worksheet.get_Range(col + row1.ToString());
                                                    Col8.Value2 = "O";// db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);
                                                }
                                            }
                                            else
                                            {

                                                Excel.Range Col0 = worksheet.get_Range(col + row1.ToString());
                                                Col0.Value2 = PV;
                                            }

                                        }
                                        catch (Exception ex) { MessageBox.Show(ex.Message); }


                                    }
                                    else
                                    {
                                        if (rd.Seq == 8)
                                        {
                                            if (db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq) != "")
                                            {
                                                if (ValueInvalid == "")
                                                    ValueInvalid = db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);
                                                else
                                                    ValueInvalid = ValueInvalid + "," + db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);

                                                Excel.Range Col8 = worksheet.get_Range(col + row1.ToString());
                                                Col8.Value2 = "O";// db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);
                                            }
                                        }
                                    }

                                }//foreach
                            }//cunt A
                        }//for
                    }

                    //if(I6=="P")
                    //{
                    //    L6 = db.QC_CheckNG(QHNo);
                    //    if(L6=="P")
                    //    {
                    //        I6 = "";
                    //    }
                    //}

                    Excel.Range Col82 = worksheet.get_Range("D17");
                    Col82.Value2 = ValueInvalid;







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
            catch (Exception ex) { MessageBox.Show("2." + ex.Message); }

        }
        public static void PrintData5601CM(string WO, string PartNo, string QCNo1)
        {
            try
            {


                string DATA = AppDomain.CurrentDomain.BaseDirectory;
                string tempPath = System.IO.Path.GetTempPath();
                string FileName = "FM-QA-056_CM.xlsx"; 
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
                int row1 = 8;
                int row2 = 9;
                int Seq = 0;
                int seq2 = 21;
                int CountRow = 0;
                string PV = "P";
                string QHNo = QCNo1;
                string FormISO = "";
                string DN = "";
                string cIssueBy1 = "";
                string cCheckBy1 = "";


                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    string Value1 = "";
                    string Value2 = "";
                    string LotNo = "";
                    string RefValue1 = "";
                    string I6 = "P";
                    string L6 = "";
                    string ValueInvalid = "";
                    string A1 = "";
                    
                    ///////////////SETValue/////////////////
                    var DValue = db.sp_46_QCSelectWO_01(WO).FirstOrDefault();
                    if (DValue != null)
                    {

                        //Sampling inspection report BRAKE CHAMBER
                        //Sampling inspection report piggyback
                        
                        DN = DValue.DayNight;
                        Excel.Range CPart = worksheet.get_Range("C3");
                        CPart.Value2 = DValue.NAME;
                        Excel.Range CStamp = worksheet.get_Range("C2");
                        CStamp.Value2 = DValue.CODE;
                        Excel.Range CName = worksheet.get_Range("C4");
                        CName.Value2 = DValue.OrderQty;

                        Excel.Range CDate = worksheet.get_Range("C5");
                        CDate.Value2 = DValue.LotNo;

                        Excel.Range CWO = worksheet.get_Range("C6");
                        CWO.Value2 = WO;


                        try
                        {
                            tb_QCHD qh = db.tb_QCHDs.Where(w => w.QCNo.Equals(QCNo1)).FirstOrDefault();
                            if (qh != null)
                            {
                                if(qh.LineName.Equals("TW02-SC_PB"))
                                {
                                    InsertToExcel(ref worksheet, "A1", "Sampling Inspection Report");
                                }
                               

                                Excel.Range App = worksheet.get_Range("I3");
                                App.Value2 = db.QC_GetUserName(qh.ApproveBy);
                                if (!qh.ApproveBy.Equals(""))
                                {
                                    Excel.Range Appdate = worksheet.get_Range("I5");
                                    Appdate.Value2 = qh.ApproveDate;
                                }

                                QHNo = qh.QCNo;
                                RefValue1 = qh.RefValue1;
                                FormISO = qh.FormISO;
                                //////////Find UserName////////////
                                var uc = db.tb_QCCheckUsers.Where(u => u.QCNo.Equals(QHNo)).ToList();


                                foreach (var rd in uc)
                                {
                                    DN = dbShowData.CheckDayN(Convert.ToDateTime(rd.ScanDate));

                                    if (rd.UDesc.Equals("Inspector"))
                                    {
                                        cIssueBy1 = rd.UserName;
                                        Excel.Range K3 = worksheet.get_Range("L3");
                                        K3.Value2 = cIssueBy1;
                                        Excel.Range K5 = worksheet.get_Range("L5");
                                        K5.Value2 = rd.ScanDate;
                                    }
                                    //if (rd.UDesc.Equals("Check By"))
                                    //{
                                    //    cCheckBy1 = rd.UserName;
                                    //    Excel.Range I3 = worksheet.get_Range("K3");
                                    //    I3.Value2 = cCheckBy1;
                                    //    Excel.Range I5 = worksheet.get_Range("K5");
                                    //    I5.Value2 = rd.ScanDate;

                                    //}
                                }
                                //Pass/Not Pass
                                if (!qh.ApproveBy.Equals(""))
                                {
                                    if (db.QC_CheckNG(QHNo) == "P")
                                    {
                                        Excel.Range L6x = worksheet.get_Range("L6");
                                        L6x.Value2 = L6;
                                    }
                                    else
                                    {
                                        Excel.Range I6x = worksheet.get_Range("I6");
                                        I6x.Value2 = I6;
                                    }
                                }

                            }

                        }
                        catch (Exception ex) { MessageBox.Show("1." + ex.Message); }




                    }

                    ////////////////////////////////////////
                    int countA = 0;
                    string col = "";
                    string col2x = "";
                    var listPoint = db.sp_46_QCSelectWO_09_QCTAGSelect(QHNo).ToList();
                    if (listPoint.Count > 0)
                    {
                        foreach (var rs in listPoint)
                        {
                            countA += 1;
                            // MessageBox.Show(countA.ToString());
                            if (countA <= 2)
                            {
                                row1 = 9;
                                col = "I";
                                col2x = "G";
                                if (countA == 2)
                                {
                                    col = "L";
                                    col2x = "H";
                                }



                                var listPart = db.tb_QCGroupParts.Where(q => q.FormISO.Equals(FormISO) && q.PartNo.Equals(DValue.CODE)).OrderBy(o => o.Seq).ToList();
                                foreach (var rd in listPart)
                                {
                                    //Start Insert Checkmark                            
                                    
                                    //if (rd.Seq <= 14)
                                    //{
                                    row1 += 1;

                                    

                                    string SetValueHD= db.get_QC_SetDataMaster2(FormISO, rd.PartNo, rd.Seq);
                                    Excel.Range SetHD = worksheet.get_Range("B" + row1.ToString());
                                    SetHD.Value2 = SetValueHD;                                   

                                    string setValueDT= db.get_QC_SetDataMaster(FormISO, rd.PartNo, rd.Seq);
                                    Excel.Range SetDT = worksheet.get_Range("D" + row1.ToString());
                                    SetDT.Value2 = setValueDT;
                                    

                                    if (setValueDT.Contains("Æ"))
                                    {                                                                               
                                        int addint = setValueDT.IndexOf("Æ");
                                        // SetDT.Characters[5, 10].Font.Color = Color.Red; // "Symbol";//AngsanaUPC
                                        SetDT.Characters[0, 2].Font.Name = "Symbol";//
                                        SetDT.Characters[2, setValueDT.Length-2].Font.Name = "Angsana New";                                     
                                    }



                                    Excel.Range SetHDA = worksheet.get_Range("A" + row1.ToString());
                                    SetHDA.Value2 = rd.Seq.ToString();
                                    

                                    //Start G=7,H=
                                    if (!rd.SetData.Equals(""))
                                    {
                                        try
                                        {
                                            var gValue = db.sp_46_QCGetValue5601(rs.BarcodeTag, QHNo, rd.Seq).FirstOrDefault();

                                            PV = "P";
                                            if (gValue.CountA > 0)
                                            {
                                                PV = "O";
                                                if (gValue.CountA == 99)
                                                    PV = "";
                                            }
                                            //if (countA == 2 && row1 == 10)
                                            //    PV = "";

                                            if (rd.SetDate2.Equals("Yes"))
                                            {
                                                PV = "";
                                                Excel.Range Col02 = worksheet.get_Range(col2x + row1.ToString());
                                                Col02.Value2 = db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);

                                                Excel.Range Col0K = worksheet.get_Range(col + row1.ToString());
                                                //Col0K.Value2 = db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);
                                                //Col0K.Font.Name = "Angsana New";// new Font("Arial", 9, FontStyle.Regular);
                                                //Col0K.Font.Size = 9;
                                                 string ValueP =  db.get_QC_DATAPointValue4(QHNo, rs.BarcodeTag, rd.Seq);

                                                if (ValueP == "OK")
                                                {
                                                    Col0K.Value2 = "P";
                                                }
                                                else if (ValueP == "NG")
                                                {
                                                    Col0K.Value2 = "O";
                                                    I6 = "";
                                                    L6 = "P";
                                                }
                                              //  Excel.Range Col0 = worksheet.get_Range(col + row1.ToString());
                                               // Col0.Value2 = PV;

                                            }
                                            else if (SetValueHD.ToUpper().Equals("OTHER"))
                                            {
                                                
                                                string RMak = db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);
                                                if (RMak != "")
                                                {
                                                    if (ValueInvalid == "")
                                                        ValueInvalid = RMak;
                                                    else
                                                        ValueInvalid = ValueInvalid + "," + RMak;

                                                    Excel.Range Col8 = worksheet.get_Range(col + row1.ToString());
                                                    Col8.Value2 = "O";// db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);

                                                    
                                                    Excel.Range Col9 = worksheet.get_Range("D21");
                                                    Col9.Value2 = ValueInvalid;// db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);
                                                }
                                            }
                                            else
                                            {

                                                Excel.Range Col0 = worksheet.get_Range(col + row1.ToString());
                                                Col0.Value2 = PV;
                                            }

                                           

                                        }
                                        catch (Exception ex) { MessageBox.Show(ex.Message); }


                                    }
                                    else
                                    {
                                        if (SetValueHD.ToUpper().Equals("OTHER"))
                                        {
                                            string RMak = db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);
                                            if (RMak != "")
                                            {
                                                if (ValueInvalid == "")
                                                    ValueInvalid = RMak;
                                                else
                                                    ValueInvalid = ValueInvalid + "," + RMak;

                                                Excel.Range Col8 = worksheet.get_Range(col + row1.ToString());
                                                Col8.Value2 = "O";// db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);
                                                Excel.Range Col9 = worksheet.get_Range("D21");
                                                Col9.Value2 = ValueInvalid;// db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);
                                            }
                                        }
                                    }

                                }//foreach
                            }//cunt A
                        }//for
                    }

                    //if(I6=="P")
                    //{
                    //    L6 = db.QC_CheckNG(QHNo);
                    //    if(L6=="P")
                    //    {
                    //        I6 = "";
                    //    }
                    //}

                  //  Excel.Range Col82 = worksheet.get_Range("D17");
                   // Col82.Value2 = ValueInvalid;







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
            catch (Exception ex) { MessageBox.Show("2." + ex.Message); }
        }
        public static void PrintData5501(string WO, string PartNo, string QCNo1)
        {
            try
            {
                //Step Report 055

                string DATA = AppDomain.CurrentDomain.BaseDirectory;
                string tempPath = System.IO.Path.GetTempPath();
                string FileName = "FM-QA-055.xlsx";
                string tempfile = tempPath + FileName;
                DATA = DATA + @"QC\" + FileName;
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    tb_ProductionHD pd = db.tb_ProductionHDs.Where(p => p.OrderNo.Equals(WO) 
                    && (p.LineName2.Equals("TW10-CB") || p.LineName2.Equals("TS10-SC_CB") || p.LineName2.Equals("TW02-SC_PB"))).FirstOrDefault();
                    if (pd != null)
                    {
                        FileName = "FM-QA-055_CM.xlsx";
                        PrintData5501CM(WO, PartNo, QCNo1);
                        return;
                    }
                }

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
                Excel.Worksheet worksheet2 = (Excel.Worksheet)sheets.get_Item(2);
                Excel.Worksheet worksheet3 = (Excel.Worksheet)sheets.get_Item(3);
                Excel.Worksheet worksheet4 = (Excel.Worksheet)sheets.get_Item(4);
                Excel.Worksheet worksheet5 = (Excel.Worksheet)sheets.get_Item(5);
                Excel.Worksheet worksheet6 = (Excel.Worksheet)sheets.get_Item(6);

                // progressBar1.Maximum = 51;
                // progressBar1.Minimum = 1;
                int row1 = 6;
                int row2 = 9;
                int Seq = 0;
                int seq2 = 21;
                int CountRow = 0;
                string PV = "P";
                string QHNo = QCNo1;
                string FormISO = "";
                int NGQ = 0;
                string DN = "";
                string ValueInvalid = "";
                string ValueInvalid2 = "";
                string ValueInvalid3 = "";
                string ValueInvalid4 = "";
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    string Value1 = "";
                    string Value2 = "";
                    string LotNo = "";
                    string RefValue1 = "";
                    string PartName = "";
                    string Remark = "";
                    bool chek24 = true;
                    decimal CKQty = 0;
                    ///////////////SETValue/////////////////
                    var DValue = db.sp_46_QCSelectWO_01(WO).FirstOrDefault();
                    if (DValue != null)
                    {
                        DN = DValue.DayNight;
                        PartName = DValue.NAME;
                        //WorkSheet1
                        Excel.Range CStamp = worksheet.get_Range("A4");
                        CStamp.Value2 = DValue.CODE;
                        Excel.Range CName = worksheet.get_Range("C4");
                        CName.Value2 = DValue.NAME;
                        Excel.Range QD = worksheet.get_Range("F4");
                        QD.Value2 = DValue.OrderQty;
                        Excel.Range CDate = worksheet.get_Range("D4");
                        CDate.Value2 = DValue.LotNo;
                        tb_QCHD qcd = db.tb_QCHDs.Where(p => p.QCNo.Equals(QCNo1)).FirstOrDefault();
                        if (qcd != null)
                        {
                            CKQty = Convert.ToDecimal(db.get_QCSumQtyTAGNG(QCNo1, "", 96));
                            Excel.Range PDQ1 = worksheet.get_Range("I4");
                            PDQ1.Value2 = CKQty;
                            Excel.Range QCOK1 = worksheet.get_Range("M4");
                            QCOK1.Value2 = qcd.OKQty;
                            Excel.Range QCNG1 = worksheet.get_Range("Q4");
                            QCNG1.Value2 = qcd.NGQty;



                            //WorkSheet2
                            Excel.Range CStamp2 = worksheet2.get_Range("A4");
                            CStamp2.Value2 = DValue.CODE;
                            Excel.Range CName2 = worksheet2.get_Range("C4");
                            CName2.Value2 = DValue.NAME;
                            Excel.Range QD2 = worksheet2.get_Range("F4");
                            QD2.Value2 = DValue.OrderQty;
                            Excel.Range CDate2 = worksheet2.get_Range("D4");
                            CDate2.Value2 = DValue.LotNo;

                            Excel.Range PDQ2 = worksheet2.get_Range("I4");
                            PDQ2.Value2 = CKQty;
                            Excel.Range QCOK2 = worksheet2.get_Range("M4");
                            QCOK2.Value2 = qcd.OKQty;
                            Excel.Range QCNG2 = worksheet2.get_Range("Q4");
                            QCNG2.Value2 = qcd.NGQty;


                            //WorkSheet3
                            Excel.Range CStamp3 = worksheet3.get_Range("A4");
                            CStamp3.Value2 = DValue.CODE;
                            Excel.Range CName3 = worksheet3.get_Range("C4");
                            CName3.Value2 = DValue.NAME;
                            Excel.Range QD3 = worksheet3.get_Range("F4");
                            QD3.Value2 = DValue.OrderQty;
                            Excel.Range CDate3 = worksheet3.get_Range("D4");
                            CDate3.Value2 = DValue.LotNo;

                            Excel.Range PDQ3 = worksheet3.get_Range("I4");
                            PDQ3.Value2 = CKQty;
                            Excel.Range QCOK3 = worksheet3.get_Range("M4");
                            QCOK3.Value2 = qcd.OKQty;
                            Excel.Range QCNG3 = worksheet3.get_Range("Q4");
                            QCNG3.Value2 = qcd.NGQty;


                            //WorkSheet4
                            Excel.Range CStamp4 = worksheet4.get_Range("A4");
                            CStamp4.Value2 = DValue.CODE;
                            Excel.Range CName4 = worksheet4.get_Range("C4");
                            CName4.Value2 = DValue.NAME;
                            Excel.Range QD4 = worksheet4.get_Range("F4");
                            QD4.Value2 = DValue.OrderQty;
                            Excel.Range CDate4 = worksheet4.get_Range("D4");
                            CDate4.Value2 = DValue.LotNo;

                            Excel.Range PDQ4 = worksheet4.get_Range("I4");
                            PDQ4.Value2 = CKQty;
                            Excel.Range QCOK4 = worksheet4.get_Range("M4");
                            QCOK4.Value2 = qcd.OKQty;
                            Excel.Range QCNG4 = worksheet4.get_Range("Q4");
                            QCNG4.Value2 = qcd.NGQty;

                            //WorkSheet5
                            Excel.Range CStamp5 = worksheet5.get_Range("A4");
                            CStamp5.Value2 = DValue.CODE;
                            Excel.Range CName5 = worksheet5.get_Range("C4");
                            CName5.Value2 = DValue.NAME;
                            Excel.Range QD5 = worksheet5.get_Range("F4");
                            QD5.Value2 = DValue.OrderQty;
                            Excel.Range CDate5 = worksheet5.get_Range("D4");
                            CDate5.Value2 = DValue.LotNo;

                            Excel.Range PDQ5 = worksheet5.get_Range("I4");
                            PDQ5.Value2 = CKQty;
                            Excel.Range QCOK5 = worksheet5.get_Range("M4");
                            QCOK5.Value2 = qcd.OKQty;
                            Excel.Range QCNG5 = worksheet5.get_Range("Q4");
                            QCNG5.Value2 = qcd.NGQty;


                            //WorkSheet6
                            Excel.Range CStamp6 = worksheet6.get_Range("A4");
                            CStamp6.Value2 = DValue.CODE;
                            Excel.Range CName6 = worksheet6.get_Range("C4");
                            CName6.Value2 = DValue.NAME;
                            Excel.Range QD6 = worksheet6.get_Range("F4");
                            QD6.Value2 = DValue.OrderQty;
                            Excel.Range CDate6 = worksheet6.get_Range("D4");
                            CDate6.Value2 = DValue.LotNo;

                            Excel.Range PDQ6 = worksheet6.get_Range("I4");
                            PDQ6.Value2 = CKQty;
                            Excel.Range QCOK6 = worksheet6.get_Range("M4");
                            QCOK6.Value2 = qcd.OKQty;
                            Excel.Range QCNG6 = worksheet6.get_Range("Q4");
                            QCNG6.Value2 = qcd.NGQty;


                        }
                        


                        chek24 = false;
                        string GP5 = "";
                        string GP6 = "";

                        //if(PartName.Contains("24"))
                        //{
                        //    chek24 = true;
                        //}
                        //if(PartName.Contains("20"))
                        //{
                        //    chek24 = true;
                        //}

                        if (PartName.Contains("30-") || PartName.Contains("-30"))
                        {
                            chek24 = false;
                            GP5 = "30-24";
                            GP6 = "D";
                            Excel.Range G19 = worksheet.get_Range("G19");
                            G19.Value2 = "P";

                            Excel.Range G192 = worksheet2.get_Range("G19");
                            G192.Value2 = "P";

                            Excel.Range G193 = worksheet3.get_Range("G19");
                            G193.Value2 = "P";

                            Excel.Range G194 = worksheet4.get_Range("G19");
                            G194.Value2 = "P";

                            Excel.Range G195 = worksheet5.get_Range("G19");
                            G195.Value2 = "P";

                            Excel.Range G196 = worksheet6.get_Range("G19");
                            G196.Value2 = "P";
                        }
                        else
                        {
                            if (PartName.Contains("16-24"))
                            {
                                GP5 = "16-24";
                                GP6 = "A";
                                Excel.Range G16 = worksheet.get_Range("G16");
                                G16.Value2 = "P";
                                Excel.Range G162 = worksheet2.get_Range("G16");
                                G162.Value2 = "P";
                                Excel.Range G163 = worksheet3.get_Range("G16");
                                G163.Value2 = "P";
                                Excel.Range G164 = worksheet4.get_Range("G16");
                                G164.Value2 = "P";
                                Excel.Range G165 = worksheet5.get_Range("G16");
                                G165.Value2 = "P";
                                Excel.Range G166 = worksheet6.get_Range("G16");
                                G166.Value2 = "P";


                            }
                            else if (PartName.Contains("20-24"))
                            {
                                GP5 = "20-24";
                                GP6 = "B";
                                Excel.Range G17 = worksheet.get_Range("G17");
                                G17.Value2 = "P";
                                Excel.Range G172 = worksheet2.get_Range("G17");
                                G172.Value2 = "P";
                                Excel.Range G173 = worksheet3.get_Range("G17");
                                G173.Value2 = "P";
                                Excel.Range G174 = worksheet4.get_Range("G17");
                                G174.Value2 = "P";
                                Excel.Range G175 = worksheet5.get_Range("G17");
                                G175.Value2 = "P";
                                Excel.Range G176 = worksheet6.get_Range("G17");
                                G176.Value2 = "P";

                            }
                            else if (PartName.Contains("24-24"))
                            {
                                GP5 = "24-24";
                                GP6 = "C";
                                Excel.Range G18 = worksheet.get_Range("G18");
                                G18.Value2 = "P";
                                Excel.Range G182 = worksheet2.get_Range("G18");
                                G182.Value2 = "P";
                                Excel.Range G183 = worksheet3.get_Range("G18");
                                G183.Value2 = "P";
                                Excel.Range G184 = worksheet4.get_Range("G18");
                                G184.Value2 = "P";
                                Excel.Range G185 = worksheet5.get_Range("G18");
                                G185.Value2 = "P";
                                Excel.Range G186 = worksheet6.get_Range("G18");
                                G186.Value2 = "P";
                            }
                        }





                        try
                        {
                            string U6 = "P";
                            string U7 = "";
                            tb_QCHD qh = db.tb_QCHDs.Where(w => w.QCNo.Equals(QCNo1)).FirstOrDefault();
                            if (qh != null)
                            {
                                FormISO = qh.FormISO;
                                Excel.Range T2 = worksheet.get_Range("T2");
                                T2.Value2 = db.QC_GetUserName(qh.ApproveBy); //qh.ApproveBy;
                                Excel.Range T22 = worksheet2.get_Range("T2");
                                T22.Value2 = db.QC_GetUserName(qh.ApproveBy); //qh.ApproveBy;
                                Excel.Range T23 = worksheet3.get_Range("T2");
                                T23.Value2 = db.QC_GetUserName(qh.ApproveBy); //qh.ApproveBy;
                                Excel.Range T24 = worksheet4.get_Range("T2");
                                T24.Value2 = db.QC_GetUserName(qh.ApproveBy); //qh.ApproveBy;
                                Excel.Range T25 = worksheet5.get_Range("T2");
                                T25.Value2 = db.QC_GetUserName(qh.ApproveBy); //qh.ApproveBy;
                                Excel.Range T26 = worksheet6.get_Range("T2");
                                T26.Value2 = db.QC_GetUserName(qh.ApproveBy); //qh.ApproveBy;

                                if (qh.ApproveBy!="")
                                {
                                    if (db.QC_CheckNG(qh.QCNo) == "P")
                                    {
                                        Excel.Range APD = worksheet.get_Range("AC4");
                                        APD.Value2 = "P";
                                        Excel.Range APD2 = worksheet2.get_Range("AC4");
                                        APD2.Value2 = "P";
                                        Excel.Range APD3 = worksheet3.get_Range("AC4");
                                        APD3.Value2 = "P";
                                        Excel.Range APD4 = worksheet4.get_Range("AC4");
                                        APD4.Value2 = "P";
                                        Excel.Range APD5 = worksheet5.get_Range("AC4");
                                        APD5.Value2 = "P";
                                        Excel.Range APD6 = worksheet6.get_Range("AC4");
                                        APD6.Value2 = "P";
                                    }
                                    else
                                    {
                                        Excel.Range APD = worksheet.get_Range("U4");
                                        APD.Value2 = "P";
                                        Excel.Range APD2 = worksheet2.get_Range("U4");
                                        APD2.Value2 = "P";
                                        Excel.Range APD3 = worksheet3.get_Range("U4");
                                        APD3.Value2 = "P";
                                        Excel.Range APD4 = worksheet4.get_Range("U4");
                                        APD4.Value2 = "P";
                                        Excel.Range APD5 = worksheet5.get_Range("U4");
                                        APD5.Value2 = "P";
                                        Excel.Range APD6 = worksheet6.get_Range("U4");
                                        APD6.Value2 = "P";
                                    }
                                }

                                if (!Convert.ToString(qh.ApproveBy).Equals(""))
                                {
                                    Excel.Range APD = worksheet.get_Range("T3");
                                    APD.Value2 = qh.ApproveDate;
                                    Excel.Range APD2 = worksheet2.get_Range("T3");
                                    APD2.Value2 = qh.ApproveDate;
                                    Excel.Range APD3 = worksheet3.get_Range("T3");
                                    APD3.Value2 = qh.ApproveDate;
                                    Excel.Range APD4 = worksheet4.get_Range("T3");
                                    APD4.Value2 = qh.ApproveDate;
                                    Excel.Range APD5 = worksheet5.get_Range("T3");
                                    APD5.Value2 = qh.ApproveDate;
                                    Excel.Range APD6 = worksheet6.get_Range("T3");
                                    APD6.Value2 = qh.ApproveDate;

                                }
                                var uc = db.tb_QCCheckUsers.Where(u => u.QCNo.Equals(QHNo)).ToList();
                                int CRow = 0;
                                foreach (var rd in uc)
                                {
                                    DN = dbShowData.CheckDayN(Convert.ToDateTime(rd.ScanDate));
                                    CRow += 1;
                                    if (rd.UDesc.Equals("Inspector"))
                                    {
                                        if (CRow == 1)
                                        {
                                            Excel.Range AH2 = worksheet.get_Range("AH2");
                                            AH2.Value2 = rd.UserName;
                                            Excel.Range AH3 = worksheet.get_Range("AH3");
                                            AH3.Value2 = rd.ScanDate;
                                            //work1
                                            Excel.Range AH22 = worksheet2.get_Range("AH2");
                                            AH22.Value2 = rd.UserName;
                                            Excel.Range AH32 = worksheet2.get_Range("AH3");
                                            AH32.Value2 = rd.ScanDate;

                                            Excel.Range AH23 = worksheet3.get_Range("AH2");
                                            AH23.Value2 = rd.UserName;
                                            Excel.Range AH33 = worksheet3.get_Range("AH3");
                                            AH33.Value2 = rd.ScanDate;

                                            Excel.Range AH24 = worksheet4.get_Range("AH2");
                                            AH24.Value2 = rd.UserName;
                                            Excel.Range AH34 = worksheet4.get_Range("AH3");
                                            AH34.Value2 = rd.ScanDate;

                                            Excel.Range AH25 = worksheet5.get_Range("AH2");
                                            AH25.Value2 = rd.UserName;
                                            Excel.Range AH35 = worksheet5.get_Range("AH3");
                                            AH35.Value2 = rd.ScanDate;

                                            Excel.Range AH26 = worksheet6.get_Range("AH2");
                                            AH26.Value2 = rd.UserName;
                                            Excel.Range AH36 = worksheet6.get_Range("AH3");
                                            AH36.Value2 = rd.ScanDate;




                                        }
                                        else if (CRow == 2)
                                        {
                                            Excel.Range AE2 = worksheet.get_Range("AE2");
                                            AE2.Value2 = rd.UserName;
                                            Excel.Range AE3 = worksheet.get_Range("AE3");
                                            AE3.Value2 = rd.ScanDate;
                                            //work1
                                            Excel.Range AE22 = worksheet2.get_Range("AE2");
                                            AE22.Value2 = rd.UserName;
                                            Excel.Range AE32 = worksheet2.get_Range("AE3");
                                            AE32.Value2 = rd.ScanDate;

                                            Excel.Range AE23 = worksheet3.get_Range("AE2");
                                            AE23.Value2 = rd.UserName;
                                            Excel.Range AE33 = worksheet3.get_Range("AE3");
                                            AE33.Value2 = rd.ScanDate;

                                            Excel.Range AE24 = worksheet4.get_Range("AE2");
                                            AE24.Value2 = rd.UserName;
                                            Excel.Range AE34 = worksheet4.get_Range("AE3");
                                            AE34.Value2 = rd.ScanDate;

                                            Excel.Range AE25 = worksheet5.get_Range("AE2");
                                            AE25.Value2 = rd.UserName;
                                            Excel.Range AE35 = worksheet5.get_Range("AE3");
                                            AE35.Value2 = rd.ScanDate;

                                            Excel.Range AE26 = worksheet6.get_Range("AE2");
                                            AE26.Value2 = rd.UserName;
                                            Excel.Range AE36 = worksheet6.get_Range("AE3");
                                            AE36.Value2 = rd.ScanDate;


                                        }
                                        else if (CRow==3)
                                        {
                                            Excel.Range AB2 = worksheet.get_Range("AB2");
                                            AB2.Value2 = rd.UserName;
                                            Excel.Range AB3 = worksheet.get_Range("AB3");
                                            AB3.Value2 = rd.ScanDate;
                                            //work1
                                            Excel.Range AB22 = worksheet2.get_Range("AB2");
                                            AB22.Value2 = rd.UserName;
                                            Excel.Range AB32 = worksheet2.get_Range("AB3");
                                            AB32.Value2 = rd.ScanDate;
                                            Excel.Range AB23 = worksheet3.get_Range("AB2");
                                            AB23.Value2 = rd.UserName;
                                            Excel.Range AB33 = worksheet3.get_Range("AB3");
                                            AB33.Value2 = rd.ScanDate;
                                            Excel.Range AB24 = worksheet4.get_Range("AB2");
                                            AB24.Value2 = rd.UserName;
                                            Excel.Range AB34 = worksheet4.get_Range("AB3");
                                            AB34.Value2 = rd.ScanDate;

                                            Excel.Range AB25 = worksheet5.get_Range("AB2");
                                            AB25.Value2 = rd.UserName;
                                            Excel.Range AB35 = worksheet5.get_Range("AB3");
                                            AB35.Value2 = rd.ScanDate;

                                            Excel.Range AB26 = worksheet6.get_Range("AB2");
                                            AB26.Value2 = rd.UserName;
                                            Excel.Range AB36 = worksheet6.get_Range("AB3");
                                            AB36.Value2 = rd.ScanDate;
                                        }
                                    }

                                    //if (rd.UDesc.Equals("Check By"))
                                    //{
                                    //    if(CRow==1)
                                    //    {
                                    //        Excel.Range X2 = worksheet.get_Range("X2");
                                    //        X2.Value2 = rd.UserName;
                                    //        Excel.Range X3 = worksheet.get_Range("X3");
                                    //        X3.Value2 = rd.ScanDate;
                                    //        //work1
                                    //        Excel.Range X22 = worksheet2.get_Range("X2");
                                    //        X22.Value2 = rd.UserName;
                                    //        Excel.Range X32 = worksheet2.get_Range("X3");
                                    //        X32.Value2 = rd.ScanDate;

                                    //        Excel.Range X23 = worksheet3.get_Range("X2");
                                    //        X23.Value2 = rd.UserName;
                                    //        Excel.Range X33 = worksheet3.get_Range("X3");
                                    //        X33.Value2 = rd.ScanDate;

                                    //        Excel.Range X24 = worksheet4.get_Range("X2");
                                    //        X24.Value2 = rd.UserName;
                                    //        Excel.Range X34 = worksheet4.get_Range("X3");
                                    //        X34.Value2 = rd.ScanDate;
                                    //    }

                                    //}
                                }

                                QHNo = qh.QCNo;
                                RefValue1 = qh.RefValue1;
                            }

                        }
                        catch { }

                    }

                    ////////////////////////////////////////
                    int SOK = 0;
                    int SNG = 0;
                    int countA = 0;
                    int TG = 0;
                    int CP = 0;
                    var listPoint = db.sp_46_QCSelectWO_09_QCTAGSelect(QHNo).ToList();
                    if (listPoint.Count > 0)
                    {
                        foreach (var rs in listPoint)
                        {
                            SOK = 0;
                            SNG = 0;
                            countA += 1;
                            TG = 0;
                            string []PPTAG = rs.ofTAG.Split('o');
                            TG = Convert.ToInt32(PPTAG[0]);
                            // MessageBox.Show(countA.ToString());
                            if (TG>0)
                            {
                                row1 = 6;
                                if (TG <= 25)
                                {
                                    CP = TG;
                                }
                                else if (TG <= 50)
                                {
                                    CP = TG - 25;
                                }
                                else if (TG <= 75)
                                {
                                    CP = TG - 50;
                                }
                                else if (TG <= 100)
                                {
                                    CP = TG - 75;
                                }
                                else if (TG <= 125)
                                {
                                    CP = TG - 100;
                                }
                                else if (TG <= 150)
                                {
                                    CP = TG - 125;
                                }

                                var listPart = db.tb_QCGroupParts.Where(q => q.FormISO.Equals(FormISO) && q.PartNo.Equals(DValue.CODE)).OrderBy(o => o.Seq).ToList();
                                foreach (var rd in listPart)
                                {
                                    //Start Insert Checkmark  
                                    row1 += 1;
                                    //Start G=7,H=
                                    if (!rd.TopPic.Equals(""))
                                    {
                                        try
                                        {
                                            Remark = "";
                                            var gValue = db.sp_46_QCGetValue5601(rs.BarcodeTag, QHNo, rd.Seq).FirstOrDefault();
                                            PV = "P";
                                            
                                            if (gValue.CountA > 0)
                                            {
                                               
                                                PV = "O";
                                                if (gValue.CountA == 99)
                                                    PV = "";
                                            }
                                            var NValue = db.sp_46_QCGetValue55501(rs.BarcodeTag, QHNo, rd.Seq).FirstOrDefault();
                                            Remark = NValue.Remark;

                                         

                                       
                                            //Excel.Range Col0 = worksheet.get_Range(Getcolumn(CP+6) + row1.ToString(), Getcolumn(CP + 6) + row1.ToString());
                                            //Col0.Value2 = PV;
                                            if (TG <= 25)
                                            {
                                                if(PV.Equals("P"))
                                                {
                                                    if(row1==15)
                                                    {
                                                        PV = "";
                                                    }
                                                }
                                                ////
                                                if (db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq) != "")
                                                {
                                                    PV = "O";
                                                    if (ValueInvalid == "")
                                                        ValueInvalid = db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);
                                                    else
                                                        ValueInvalid = ValueInvalid + "," + db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);
                                                    
                                                }
                                                Excel.Range Col0 = worksheet.get_Range(Getcolumn(CP + 6) + row1.ToString(), Getcolumn(CP + 6) + row1.ToString());
                                                Col0.Value2 = PV;
                                                
                                            }
                                            else if (TG <= 50)
                                            {
                                                if (PV.Equals("P"))
                                                {
                                                    if (row1 == 15)
                                                    {
                                                        PV = "";
                                                    }
                                                }
                                                ////
                                                if (db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq) != "")
                                                {
                                                    PV = "O";
                                                    if (ValueInvalid2 == "")
                                                        ValueInvalid2 = db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);
                                                    else
                                                        ValueInvalid2 = ValueInvalid + "," + db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);
                                                }

                                                Excel.Range Col02 = worksheet2.get_Range(Getcolumn(CP + 6) + row1.ToString(), Getcolumn(CP + 6) + row1.ToString());
                                                Col02.Value2 = PV;
                                               
                                            }
                                            else if (TG <= 75)
                                            {
                                                if (PV.Equals("P"))
                                                {
                                                    if (row1 == 15)
                                                    {
                                                        PV = "";
                                                    }
                                                }
                                                ////
                                                if (db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq) != "")
                                                {
                                                    PV = "O";
                                                    if (ValueInvalid3 == "")
                                                        ValueInvalid3 = db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);
                                                    else
                                                        ValueInvalid3 = ValueInvalid + "," + db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);
                                                }

                                                Excel.Range Col03 = worksheet3.get_Range(Getcolumn(CP + 6) + row1.ToString(), Getcolumn(CP + 6) + row1.ToString());
                                                Col03.Value2 = PV;
                                                
                                            }
                                            else if (TG <= 100)
                                            {
                                                if (PV.Equals("P"))
                                                {
                                                    if (row1 == 15)
                                                    {
                                                        PV = "";
                                                    }
                                                }
                                                ////
                                                if (db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq) != "")
                                                {
                                                    PV = "O";
                                                    if (ValueInvalid4 == "")
                                                        ValueInvalid4 = db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);
                                                    else
                                                        ValueInvalid4 = ValueInvalid + "," + db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);
                                                }

                                                Excel.Range Col04 = worksheet4.get_Range(Getcolumn(CP + 6) + row1.ToString(), Getcolumn(CP + 6) + row1.ToString());
                                                Col04.Value2 = PV;
                                                
                                            }
                                            else if (TG <= 125)
                                            {
                                                if (PV.Equals("P"))
                                                {
                                                    if (row1 == 15)
                                                    {
                                                        PV = "";
                                                    }
                                                }
                                                ////
                                                if (db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq) != "")
                                                {
                                                    PV = "O";
                                                    if (ValueInvalid4 == "")
                                                        ValueInvalid4 = db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);
                                                    else
                                                        ValueInvalid4 = ValueInvalid + "," + db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);
                                                }

                                                Excel.Range Col05 = worksheet5.get_Range(Getcolumn(CP + 6) + row1.ToString(), Getcolumn(CP + 6) + row1.ToString());
                                                Col05.Value2 = PV;

                                            }
                                            else if (TG <= 150)
                                            {
                                                if (PV.Equals("P"))
                                                {
                                                    if (row1 == 15)
                                                    {
                                                        PV = "";
                                                    }
                                                }
                                                ////
                                                if (db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq) != "")
                                                {
                                                    PV = "O";
                                                    if (ValueInvalid4 == "")
                                                        ValueInvalid4 = db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);
                                                    else
                                                        ValueInvalid4 = ValueInvalid + "," + db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);
                                                }

                                                Excel.Range Col06 = worksheet6.get_Range(Getcolumn(CP + 6) + row1.ToString(), Getcolumn(CP + 6) + row1.ToString());
                                                Col06.Value2 = PV;

                                            }

                                            if (!Remark.Equals(""))
                                            {
                                                if (TG <= 25)
                                                {
                                                    Excel.Range Col1 = worksheet.get_Range("AF" + Convert.ToString(row1));
                                                    Col1.Value2 = Remark;
                                                }
                                                else if (TG <= 50)
                                                {

                                                    Excel.Range Col12 = worksheet2.get_Range("AF" + Convert.ToString(row1));
                                                    Col12.Value2 = Remark;
                                                }
                                                else if (TG <= 75)
                                                {

                                                    Excel.Range Col13 = worksheet3.get_Range("AF" + Convert.ToString(row1));
                                                    Col13.Value2 = Remark;
                                                }
                                                else if (TG <= 100)
                                                {
                                                    Excel.Range Col14 = worksheet4.get_Range("AF" + Convert.ToString(row1));
                                                    Col14.Value2 = Remark;
                                                }
                                                else if (TG <= 125)
                                                {
                                                    Excel.Range Col15 = worksheet5.get_Range("AF" + Convert.ToString(row1));
                                                    Col15.Value2 = Remark;
                                                }
                                                else if (TG <= 150)
                                                {
                                                    Excel.Range Col16 = worksheet6.get_Range("AF" + Convert.ToString(row1));
                                                    Col16.Value2 = Remark;
                                                }
                                            }                                                                                            

                                        }
                                        catch (Exception ex) { MessageBox.Show(ex.Message); }
                                        //}




                                    }
                                    //SumNG//
                                }//foreach
                            }//cunt A

                            //Find count Tag
                            if (TG <= 25)
                            {
                                Excel.Range GNG = worksheet.get_Range(Getcolumn(CP + 6) + "21");
                                GNG.Value2 = db.get_QCSumQtyTAGNG(QHNo, rs.BarcodeTag, 3);
                                Excel.Range GOK = worksheet.get_Range(Getcolumn(CP + 6) + "20");
                                GOK.Value2 = db.get_QCSumQtyTAGNG(QHNo, rs.BarcodeTag, 4);

                                Excel.Range C15 = worksheet.get_Range("C15");
                                C15.Value2 = ValueInvalid;
                            }
                            else if (TG <= 50)
                            {
                                Excel.Range GNG2 = worksheet2.get_Range(Getcolumn(CP + 6) + "21");
                                GNG2.Value2 = db.get_QCSumQtyTAGNG(QHNo, rs.BarcodeTag, 3);
                                Excel.Range GOK2 = worksheet2.get_Range(Getcolumn(CP + 6) + "20");
                                GOK2.Value2 = db.get_QCSumQtyTAGNG(QHNo, rs.BarcodeTag, 4);
                                Excel.Range C15 = worksheet2.get_Range("C15");
                                C15.Value2 = ValueInvalid2;
                            }
                            else if (TG <= 75)
                            {
                                Excel.Range GNG3 = worksheet3.get_Range(Getcolumn(CP + 6) + "21");
                                GNG3.Value2 = db.get_QCSumQtyTAGNG(QHNo, rs.BarcodeTag, 3);
                                Excel.Range GOK3 = worksheet3.get_Range(Getcolumn(CP + 6) + "20");
                                GOK3.Value2 = db.get_QCSumQtyTAGNG(QHNo, rs.BarcodeTag, 4);
                                Excel.Range C15 = worksheet3.get_Range("C15");
                                C15.Value2 = ValueInvalid3;
                            }
                            else if (TG <= 100)
                            {
                                Excel.Range GNG4 = worksheet4.get_Range(Getcolumn(CP + 6) + "21");
                                GNG4.Value2 = db.get_QCSumQtyTAGNG(QHNo, rs.BarcodeTag, 3);
                                Excel.Range GOK4 = worksheet4.get_Range(Getcolumn(CP + 6) + "20");
                                GOK4.Value2 = db.get_QCSumQtyTAGNG(QHNo, rs.BarcodeTag, 4);
                                Excel.Range C15 = worksheet4.get_Range("C15");
                                C15.Value2 = ValueInvalid4;
                            }
                            else if (TG <= 125)
                            {
                                Excel.Range GNG5 = worksheet5.get_Range(Getcolumn(CP + 6) + "21");
                                GNG5.Value2 = db.get_QCSumQtyTAGNG(QHNo, rs.BarcodeTag, 3);
                                Excel.Range GOK5 = worksheet5.get_Range(Getcolumn(CP + 6) + "20");
                                GOK5.Value2 = db.get_QCSumQtyTAGNG(QHNo, rs.BarcodeTag, 4);
                                Excel.Range C15 = worksheet5.get_Range("C15");
                                C15.Value2 = ValueInvalid4;
                            }
                            else if (TG <= 150)
                            {
                                Excel.Range GNG6 = worksheet6.get_Range(Getcolumn(CP + 6) + "21");
                                GNG6.Value2 = db.get_QCSumQtyTAGNG(QHNo, rs.BarcodeTag, 3);
                                Excel.Range GOK6 = worksheet6.get_Range(Getcolumn(CP + 6) + "20");
                                GOK6.Value2 = db.get_QCSumQtyTAGNG(QHNo, rs.BarcodeTag, 4);
                                Excel.Range C15 = worksheet6.get_Range("C15");
                                C15.Value2 = ValueInvalid4;
                            }
                        }//for
                    }



                }

                excelBook.SaveAs(tempfile);
                excelBook.Close(false);
                excelApp.Quit();

                releaseObject(worksheet);
                releaseObject(worksheet2);
                releaseObject(worksheet3);
                releaseObject(worksheet4);
                releaseObject(worksheet5);
                releaseObject(worksheet6);
                releaseObject(excelBook);
                releaseObject(excelApp);
                Marshal.FinalReleaseComObject(worksheet);
                Marshal.FinalReleaseComObject(worksheet2);
                Marshal.FinalReleaseComObject(worksheet3);
                Marshal.FinalReleaseComObject(worksheet4);
                Marshal.FinalReleaseComObject(worksheet5);
                Marshal.FinalReleaseComObject(worksheet6);
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

        }
        public static void PrintData5501CM(string WO, string PartNo, string QCNo1)
        {
            try
            {
                //Step Report 055
              //  MessageBox.Show("eeed");

                string DATA = AppDomain.CurrentDomain.BaseDirectory;
                string tempPath = System.IO.Path.GetTempPath();
                string FileName = "FM-QA-055CM.xlsx";
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
                Excel.Worksheet worksheet2 = (Excel.Worksheet)sheets.get_Item(2);
                Excel.Worksheet worksheet3 = (Excel.Worksheet)sheets.get_Item(3);
                Excel.Worksheet worksheet4 = (Excel.Worksheet)sheets.get_Item(4);
                Excel.Worksheet worksheet5 = (Excel.Worksheet)sheets.get_Item(5);
                Excel.Worksheet worksheet6 = (Excel.Worksheet)sheets.get_Item(6);

                // progressBar1.Maximum = 51;
                // progressBar1.Minimum = 1;
                int row1 = 6;
                int row2 = 9;
                int Seq = 0;
                int seq2 = 21;
                int CountRow = 0;
                string PV = "P";
                string QHNo = QCNo1;
                string FormISO = "";
                int NGQ = 0;
                string DN = "";
                string ValueInvalid = "";
                string ValueInvalid2 = "";
                string ValueInvalid3 = "";
                string ValueInvalid4 = "";
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    string Value1 = "";
                    string Value2 = "";
                    string LotNo = "";
                    string RefValue1 = "";
                    string PartName = "";
                    string Remark = "";
                    bool chek24 = true;
                    decimal CKQty = 0;
                    int PackI = 6;
                    decimal SNP = 0;
                    ///////////////SETValue/////////////////
                    var DValue = db.sp_46_QCSelectWO_01(WO).FirstOrDefault();
                    if (DValue != null)
                    {
                        DN = DValue.DayNight;
                        PartName = DValue.NAME;
                        //WorkSheet1
                        Excel.Range CStamp = worksheet.get_Range("A4");
                        CStamp.Value2 = DValue.CODE;
                        Excel.Range CName = worksheet.get_Range("C4");
                        CName.Value2 = DValue.NAME;
                        Excel.Range QD = worksheet.get_Range("F4");
                        QD.Value2 = DValue.OrderQty;
                        Excel.Range CDate = worksheet.get_Range("D4");
                        CDate.Value2 = DValue.LotNo;
                        // db.sp_46_QCSelectWO_12_UpdateOKQty(QCNo1);
                        string AA = "CHAMBER       100% INSPECTION CHECK SHEET";

                        tb_QCHD qcd = db.tb_QCHDs.Where(p => p.QCNo.Equals(QHNo)).FirstOrDefault();
                        if (qcd != null)
                        {
                            //
                            if (qcd.LineName.Equals(""))
                            {
                                AA = " PIGGYBACK            100% INSPECTION CHECK SHEET";
                            }
                           
                            SNP = Convert.ToDecimal(db.getLOtSizeTPICS_Dynamics(qcd.PartNo));
                            //
                            CKQty = Convert.ToDecimal(db.get_QCSumQtyTAGNG(QCNo1, "", 96));
                            Excel.Range PDQ1 = worksheet.get_Range("I4");
                            PDQ1.Value2 = CKQty;
                            //Excel.Range QCOK1 = worksheet.get_Range("M4");
                            //QCOK1.Value2 = qcd.OKQty;
                            Excel.Range QCNG1 = worksheet.get_Range("Q4");
                            QCNG1.Value2 = qcd.NGQty;                            
                            Excel.Range G15SNP1 = worksheet.get_Range("F20");
                            G15SNP1.Value2 = SNP;
                            InsertToExcel(ref worksheet, "A1", AA);
                            //tb_QCTAG tcg = db.tb_QCTAGs.Where(p => p.QCNo.Equals(QCNo1)).FirstOrDefault();
                            //if(tcg!=null)
                            //{
                            //    string[] TAG = tcg.BarcodeTag.Split(',');
                            //    if(TAG[2]=="4")
                            //    {
                            //        PackI = 4;
                            //    }
                            //}



                            //WorkSheet2
                            Excel.Range CStamp2 = worksheet2.get_Range("A4");
                            CStamp2.Value2 = DValue.CODE;
                            Excel.Range CName2 = worksheet2.get_Range("C4");
                            CName2.Value2 = DValue.NAME;
                            Excel.Range QD2 = worksheet2.get_Range("F4");
                            QD2.Value2 = DValue.OrderQty;
                            Excel.Range CDate2 = worksheet2.get_Range("D4");
                            CDate2.Value2 = DValue.LotNo;
                            Excel.Range PDQ2 = worksheet2.get_Range("I4");
                            PDQ2.Value2 = CKQty;
                            //Excel.Range QCOK2 = worksheet2.get_Range("M4");
                            //QCOK2.Value2 = qcd.OKQty;
                            Excel.Range QCNG2 = worksheet2.get_Range("Q4");
                            QCNG2.Value2 = qcd.NGQty;
                            Excel.Range G15SNP2 = worksheet2.get_Range("F20");
                            G15SNP2.Value2 = SNP;
                            InsertToExcel(ref worksheet2, "A1", AA);

                            //WorkSheet3
                            Excel.Range CStamp3 = worksheet3.get_Range("A4");
                            CStamp3.Value2 = DValue.CODE;
                            Excel.Range CName3 = worksheet3.get_Range("C4");
                            CName3.Value2 = DValue.NAME;
                            Excel.Range QD3 = worksheet3.get_Range("F4");
                            QD3.Value2 = DValue.OrderQty;
                            Excel.Range CDate3 = worksheet3.get_Range("D4");
                            CDate3.Value2 = DValue.LotNo;
                            Excel.Range PDQ3 = worksheet3.get_Range("I4");
                            PDQ3.Value2 = CKQty;
                            //Excel.Range QCOK3 = worksheet3.get_Range("M4");
                            //QCOK3.Value2 = qcd.OKQty;
                            Excel.Range QCNG3 = worksheet3.get_Range("Q4");
                            QCNG3.Value2 = qcd.NGQty;
                            Excel.Range G15SNP3 = worksheet3.get_Range("F20");
                            G15SNP3.Value2 = SNP;
                            InsertToExcel(ref worksheet3, "A1", AA);


                            //WorkSheet4
                            Excel.Range CStamp4 = worksheet4.get_Range("A4");
                            CStamp4.Value2 = DValue.CODE;
                            Excel.Range CName4 = worksheet4.get_Range("C4");
                            CName4.Value2 = DValue.NAME;
                            Excel.Range QD4 = worksheet4.get_Range("F4");
                            QD4.Value2 = DValue.OrderQty;
                            Excel.Range CDate4 = worksheet4.get_Range("D4");
                            CDate4.Value2 = DValue.LotNo;
                            Excel.Range PDQ4 = worksheet4.get_Range("I4");
                            PDQ4.Value2 = CKQty;
                            //Excel.Range QCOK4 = worksheet4.get_Range("M4");
                            //QCOK4.Value2 = qcd.OKQty;
                            Excel.Range QCNG4 = worksheet4.get_Range("Q4");
                            QCNG4.Value2 = qcd.NGQty;
                            Excel.Range G15SNP4 = worksheet4.get_Range("F20");
                            G15SNP4.Value2 = SNP;
                            InsertToExcel(ref worksheet4, "A1", AA);

                            //WorkSheet5
                            Excel.Range CStamp5 = worksheet5.get_Range("A4");
                            CStamp5.Value2 = DValue.CODE;
                            Excel.Range CName5 = worksheet5.get_Range("C4");
                            CName5.Value2 = DValue.NAME;
                            Excel.Range QD5 = worksheet5.get_Range("F4");
                            QD5.Value2 = DValue.OrderQty;
                            Excel.Range CDate5 = worksheet5.get_Range("D4");
                            CDate5.Value2 = DValue.LotNo;
                            Excel.Range PDQ5 = worksheet5.get_Range("I4");
                            PDQ5.Value2 = CKQty;
                            //Excel.Range QCOK5 = worksheet5.get_Range("M4");
                            //QCOK5.Value2 = qcd.OKQty;
                            Excel.Range QCNG5 = worksheet5.get_Range("Q4");
                            QCNG5.Value2 = qcd.NGQty;
                            Excel.Range G15SNP5 = worksheet5.get_Range("F20");
                            G15SNP5.Value2 = SNP;
                            InsertToExcel(ref worksheet5, "A1", AA);


                            //WorkSheet6
                            Excel.Range CStamp6 = worksheet6.get_Range("A4");
                            CStamp6.Value2 = DValue.CODE;
                            Excel.Range CName6 = worksheet6.get_Range("C4");
                            CName6.Value2 = DValue.NAME;
                            Excel.Range QD6 = worksheet6.get_Range("F4");
                            QD6.Value2 = DValue.OrderQty;
                            Excel.Range CDate6 = worksheet6.get_Range("D4");
                            CDate6.Value2 = DValue.LotNo;
                            Excel.Range PDQ6 = worksheet6.get_Range("I4");
                            PDQ6.Value2 = CKQty;
                            //Excel.Range QCOK6 = worksheet6.get_Range("M4");
                            //QCOK6.Value2 = qcd.OKQty;
                            Excel.Range QCNG6 = worksheet6.get_Range("Q4");
                            QCNG6.Value2 = qcd.NGQty;
                            Excel.Range G15SNP6 = worksheet6.get_Range("F20");
                            G15SNP6.Value2 = SNP;
                            InsertToExcel(ref worksheet6, "A1", AA);


                        }
                        

                        chek24 = false;
                      //  string GP5 = "";
                       // string GP6 = "";


                        
                        
                        if (PackI==6)
                            {
                               // GP5 = "16-24";
                               // GP6 = "A";

                                //Excel.Range G16 = worksheet.get_Range("G19");
                                //G16.Value2 = "P";
                                //Excel.Range G162 = worksheet2.get_Range("G19");
                                //G162.Value2 = "P";
                                //Excel.Range G163 = worksheet3.get_Range("G19");
                                //G163.Value2 = "P";
                                //Excel.Range G164 = worksheet4.get_Range("G19");
                                //G164.Value2 = "P";
                                //Excel.Range G165 = worksheet5.get_Range("G19");
                                //G165.Value2 = "P";
                                //Excel.Range G166 = worksheet6.get_Range("G19");
                                //G166.Value2 = "P";


                            }
                            else if (PackI==4)
                            {
                               // GP5 = "20-24";
                               // GP6 = "B";

                                //Excel.Range G17 = worksheet.get_Range("G20");
                                //G17.Value2 = "P";
                                //Excel.Range G172 = worksheet2.get_Range("G20");
                                //G172.Value2 = "P";
                                //Excel.Range G173 = worksheet3.get_Range("G20");
                                //G173.Value2 = "P";
                                //Excel.Range G174 = worksheet4.get_Range("G20");
                                //G174.Value2 = "P";
                                //Excel.Range G175 = worksheet5.get_Range("G20");
                                //G175.Value2 = "P";
                                //Excel.Range G176 = worksheet6.get_Range("G20");
                                //G176.Value2 = "P";

                            }
                            
                        





                        try
                        {
                            string U6 = "P";
                            string U7 = "";
                            tb_QCHD qh = db.tb_QCHDs.Where(w => w.QCNo.Equals(QCNo1)).FirstOrDefault();
                            if (qh != null)
                            {
                                FormISO = qh.FormISO;
                                Excel.Range T2 = worksheet.get_Range("T2");
                                T2.Value2 = db.QC_GetUserName(qh.ApproveBy); //qh.ApproveBy;
                                Excel.Range T22 = worksheet2.get_Range("T2");
                                T22.Value2 = db.QC_GetUserName(qh.ApproveBy); //qh.ApproveBy;
                                Excel.Range T23 = worksheet3.get_Range("T2");
                                T23.Value2 = db.QC_GetUserName(qh.ApproveBy); //qh.ApproveBy;
                                Excel.Range T24 = worksheet4.get_Range("T2");
                                T24.Value2 = db.QC_GetUserName(qh.ApproveBy); //qh.ApproveBy;
                                Excel.Range T25 = worksheet5.get_Range("T2");
                                T25.Value2 = db.QC_GetUserName(qh.ApproveBy); //qh.ApproveBy;
                                Excel.Range T26 = worksheet6.get_Range("T2");
                                T26.Value2 = db.QC_GetUserName(qh.ApproveBy); //qh.ApproveBy;

                                if (qh.ApproveBy != "")
                                {
                                    if (db.QC_CheckNG(qh.QCNo) == "P")
                                    {
                                        Excel.Range APD = worksheet.get_Range("AC4");
                                        APD.Value2 = "P";
                                        Excel.Range APD2 = worksheet2.get_Range("AC4");
                                        APD2.Value2 = "P";
                                        Excel.Range APD3 = worksheet3.get_Range("AC4");
                                        APD3.Value2 = "P";
                                        Excel.Range APD4 = worksheet4.get_Range("AC4");
                                        APD4.Value2 = "P";
                                        Excel.Range APD5 = worksheet5.get_Range("AC4");
                                        APD5.Value2 = "P";
                                        Excel.Range APD6 = worksheet6.get_Range("AC4");
                                        APD6.Value2 = "P";
                                    }
                                    else
                                    {
                                        Excel.Range APD = worksheet.get_Range("U4");
                                        APD.Value2 = "P";
                                        Excel.Range APD2 = worksheet2.get_Range("U4");
                                        APD2.Value2 = "P";
                                        Excel.Range APD3 = worksheet3.get_Range("U4");
                                        APD3.Value2 = "P";
                                        Excel.Range APD4 = worksheet4.get_Range("U4");
                                        APD4.Value2 = "P";
                                        Excel.Range APD5 = worksheet5.get_Range("U4");
                                        APD5.Value2 = "P";
                                        Excel.Range APD6 = worksheet6.get_Range("U4");
                                        APD6.Value2 = "P";
                                    }
                                }

                                if (!Convert.ToString(qh.ApproveBy).Equals(""))
                                {
                                    Excel.Range APD = worksheet.get_Range("T3");
                                    APD.Value2 = qh.ApproveDate;
                                    Excel.Range APD2 = worksheet2.get_Range("T3");
                                    APD2.Value2 = qh.ApproveDate;
                                    Excel.Range APD3 = worksheet3.get_Range("T3");
                                    APD3.Value2 = qh.ApproveDate;
                                    Excel.Range APD4 = worksheet4.get_Range("T3");
                                    APD4.Value2 = qh.ApproveDate;
                                    Excel.Range APD5 = worksheet5.get_Range("T3");
                                    APD5.Value2 = qh.ApproveDate;
                                    Excel.Range APD6 = worksheet6.get_Range("T3");
                                    APD6.Value2 = qh.ApproveDate;

                                }
                                var uc = db.tb_QCCheckUsers.Where(u => u.QCNo.Equals(QCNo1)).ToList();
                                int CRow = 0;
                                foreach (var rd in uc)
                                {
                                    DN = dbShowData.CheckDayN(Convert.ToDateTime(rd.ScanDate));
                                    CRow += 1;
                                    if (rd.UDesc.Equals("Inspector"))
                                    {
                                        if (CRow == 1)
                                        {
                                            Excel.Range AH2 = worksheet.get_Range("AH2");
                                            AH2.Value2 = rd.UserName;
                                            Excel.Range AH3 = worksheet.get_Range("AH3");
                                            AH3.Value2 = rd.ScanDate;
                                            //work1
                                            Excel.Range AH22 = worksheet2.get_Range("AH2");
                                            AH22.Value2 = rd.UserName;
                                            Excel.Range AH32 = worksheet2.get_Range("AH3");
                                            AH32.Value2 = rd.ScanDate;

                                            Excel.Range AH23 = worksheet3.get_Range("AH2");
                                            AH23.Value2 = rd.UserName;
                                            Excel.Range AH33 = worksheet3.get_Range("AH3");
                                            AH33.Value2 = rd.ScanDate;

                                            Excel.Range AH24 = worksheet4.get_Range("AH2");
                                            AH24.Value2 = rd.UserName;
                                            Excel.Range AH34 = worksheet4.get_Range("AH3");
                                            AH34.Value2 = rd.ScanDate;

                                            Excel.Range AH25 = worksheet5.get_Range("AH2");
                                            AH25.Value2 = rd.UserName;
                                            Excel.Range AH35 = worksheet5.get_Range("AH3");
                                            AH35.Value2 = rd.ScanDate;

                                            Excel.Range AH26 = worksheet6.get_Range("AH2");
                                            AH26.Value2 = rd.UserName;
                                            Excel.Range AH36 = worksheet6.get_Range("AH3");
                                            AH36.Value2 = rd.ScanDate;




                                        }
                                        else if (CRow == 2)
                                        {
                                            Excel.Range AE2 = worksheet.get_Range("AE2");
                                            AE2.Value2 = rd.UserName;
                                            Excel.Range AE3 = worksheet.get_Range("AE3");
                                            AE3.Value2 = rd.ScanDate;
                                            //work1
                                            Excel.Range AE22 = worksheet2.get_Range("AE2");
                                            AE22.Value2 = rd.UserName;
                                            Excel.Range AE32 = worksheet2.get_Range("AE3");
                                            AE32.Value2 = rd.ScanDate;

                                            Excel.Range AE23 = worksheet3.get_Range("AE2");
                                            AE23.Value2 = rd.UserName;
                                            Excel.Range AE33 = worksheet3.get_Range("AE3");
                                            AE33.Value2 = rd.ScanDate;

                                            Excel.Range AE24 = worksheet4.get_Range("AE2");
                                            AE24.Value2 = rd.UserName;
                                            Excel.Range AE34 = worksheet4.get_Range("AE3");
                                            AE34.Value2 = rd.ScanDate;

                                            Excel.Range AE25 = worksheet5.get_Range("AE2");
                                            AE25.Value2 = rd.UserName;
                                            Excel.Range AE35 = worksheet5.get_Range("AE3");
                                            AE35.Value2 = rd.ScanDate;

                                            Excel.Range AE26 = worksheet6.get_Range("AE2");
                                            AE26.Value2 = rd.UserName;
                                            Excel.Range AE36 = worksheet6.get_Range("AE3");
                                            AE36.Value2 = rd.ScanDate;


                                        }
                                        else if (CRow == 3)
                                        {
                                            Excel.Range AB2 = worksheet.get_Range("AB2");
                                            AB2.Value2 = rd.UserName;
                                            Excel.Range AB3 = worksheet.get_Range("AB3");
                                            AB3.Value2 = rd.ScanDate;
                                            //work1
                                            Excel.Range AB22 = worksheet2.get_Range("AB2");
                                            AB22.Value2 = rd.UserName;
                                            Excel.Range AB32 = worksheet2.get_Range("AB3");
                                            AB32.Value2 = rd.ScanDate;
                                            Excel.Range AB23 = worksheet3.get_Range("AB2");
                                            AB23.Value2 = rd.UserName;
                                            Excel.Range AB33 = worksheet3.get_Range("AB3");
                                            AB33.Value2 = rd.ScanDate;
                                            Excel.Range AB24 = worksheet4.get_Range("AB2");
                                            AB24.Value2 = rd.UserName;
                                            Excel.Range AB34 = worksheet4.get_Range("AB3");
                                            AB34.Value2 = rd.ScanDate;

                                            Excel.Range AB25 = worksheet5.get_Range("AB2");
                                            AB25.Value2 = rd.UserName;
                                            Excel.Range AB35 = worksheet5.get_Range("AB3");
                                            AB35.Value2 = rd.ScanDate;

                                            Excel.Range AB26 = worksheet6.get_Range("AB2");
                                            AB26.Value2 = rd.UserName;
                                            Excel.Range AB36 = worksheet6.get_Range("AB3");
                                            AB36.Value2 = rd.ScanDate;
                                        }
                                    }

                                    //if (rd.UDesc.Equals("Check By"))
                                    //{
                                    //    if(CRow==1)
                                    //    {
                                    //        Excel.Range X2 = worksheet.get_Range("X2");
                                    //        X2.Value2 = rd.UserName;
                                    //        Excel.Range X3 = worksheet.get_Range("X3");
                                    //        X3.Value2 = rd.ScanDate;
                                    //        //work1
                                    //        Excel.Range X22 = worksheet2.get_Range("X2");
                                    //        X22.Value2 = rd.UserName;
                                    //        Excel.Range X32 = worksheet2.get_Range("X3");
                                    //        X32.Value2 = rd.ScanDate;

                                    //        Excel.Range X23 = worksheet3.get_Range("X2");
                                    //        X23.Value2 = rd.UserName;
                                    //        Excel.Range X33 = worksheet3.get_Range("X3");
                                    //        X33.Value2 = rd.ScanDate;

                                    //        Excel.Range X24 = worksheet4.get_Range("X2");
                                    //        X24.Value2 = rd.UserName;
                                    //        Excel.Range X34 = worksheet4.get_Range("X3");
                                    //        X34.Value2 = rd.ScanDate;
                                    //    }

                                    //}
                                }

                                QHNo = qh.QCNo;
                                RefValue1 = qh.RefValue1;
                            }

                        }
                        catch { }

                    }

                    ////////////////////////////////////////

                    //Insert Header//
                    int rx = 6;
                    int TGx = 0;
                    int RowCountA = 0;
                    string Rank = "B";
                    var listPoint2 = db.tb_QCTAGs.Where(p => p.QCNo.Equals(QCNo1)).FirstOrDefault();
                    if (listPoint2 != null)
                    {
                        // 1of10
                        string[] PPTAG2 = listPoint2.ofTAG.Split('f');
                        if (PPTAG2.Length > 1)
                            TGx = Convert.ToInt32(PPTAG2[1]);
                    }
                    var listPart2 = db.tb_QCGroupParts.Where(q => q.FormISO.Equals(FormISO) && q.PartNo.Equals(DValue.CODE)).OrderBy(o => o.Seq).ToList();
                    foreach (var rd in listPart2)
                    {
                        rx += 1;
                        RowCountA += 1;
                        if (RowCountA == 1)
                            Rank = "A";
                        else
                            Rank = "B";

                        string SetValueHD = db.get_QC_SetDataMaster2(FormISO, rd.PartNo, rd.Seq);
                        string setValueDT = db.get_QC_SetDataMaster(FormISO, rd.PartNo, rd.Seq);
                       
                        InsertToExcel(ref worksheet, "B" + rx.ToString(), SetValueHD);
                        InsertToExcel(ref worksheet, "C" + rx.ToString(), setValueDT);

                        InsertToExcel(ref worksheet, "A" + rx.ToString(), RowCountA.ToString());
                        InsertToExcel(ref worksheet, "D" + rx.ToString(), "Visual");
                        InsertToExcel(ref worksheet, "E" + rx.ToString(), Rank);
                        InsertToExcel(ref worksheet, "F" + rx.ToString(), "100%");

                        if (TGx > 25)
                        {
                            InsertToExcel(ref worksheet2, "B" + rx.ToString(), SetValueHD);
                            InsertToExcel(ref worksheet2, "C" + rx.ToString(), setValueDT);

                            InsertToExcel(ref worksheet2, "A" + rx.ToString(), RowCountA.ToString());
                            InsertToExcel(ref worksheet2, "D" + rx.ToString(), "Visual");
                            InsertToExcel(ref worksheet2, "E" + rx.ToString(), Rank);
                            InsertToExcel(ref worksheet2, "F" + rx.ToString(), "100%");
                        }
                        if (TGx > 50)
                        {
                            InsertToExcel(ref worksheet3, "B" + rx.ToString(), SetValueHD);
                            InsertToExcel(ref worksheet3, "C" + rx.ToString(), setValueDT);

                            InsertToExcel(ref worksheet3, "A" + rx.ToString(), RowCountA.ToString());
                            InsertToExcel(ref worksheet3, "D" + rx.ToString(), "Visual");
                            InsertToExcel(ref worksheet3, "E" + rx.ToString(), Rank);
                            InsertToExcel(ref worksheet3, "F" + rx.ToString(), "100%");
                        }
                        if (TGx > 75)
                        {
                            InsertToExcel(ref worksheet4, "B" + rx.ToString(), SetValueHD);
                            InsertToExcel(ref worksheet4, "C" + rx.ToString(), setValueDT);

                            InsertToExcel(ref worksheet4, "A" + rx.ToString(), RowCountA.ToString());
                            InsertToExcel(ref worksheet4, "D" + rx.ToString(), "Visual");
                            InsertToExcel(ref worksheet4, "E" + rx.ToString(), Rank);
                            InsertToExcel(ref worksheet4, "F" + rx.ToString(), "100%");
                        }
                        if (TGx > 100)
                        {
                            InsertToExcel(ref worksheet5, "B" + rx.ToString(), SetValueHD);
                            InsertToExcel(ref worksheet5, "C" + rx.ToString(), setValueDT);

                            InsertToExcel(ref worksheet5, "A" + rx.ToString(), RowCountA.ToString());
                            InsertToExcel(ref worksheet5, "D" + rx.ToString(), "Visual");
                            InsertToExcel(ref worksheet5, "E" + rx.ToString(), Rank);
                            InsertToExcel(ref worksheet5, "F" + rx.ToString(), "100%");
                        }
                        if (TGx > 125)
                        {
                            InsertToExcel(ref worksheet6, "B" + rx.ToString(), SetValueHD);
                            InsertToExcel(ref worksheet6, "C" + rx.ToString(), setValueDT);

                            InsertToExcel(ref worksheet6, "A" + rx.ToString(), RowCountA.ToString());
                            InsertToExcel(ref worksheet6, "D" + rx.ToString(), "Visual");
                            InsertToExcel(ref worksheet6, "E" + rx.ToString(), Rank);
                            InsertToExcel(ref worksheet6, "F" + rx.ToString(), "100%");
                        }



                    }


                    ///////////////////////////////////////
                    int SOK = 0;
                    int SNG = 0;
                    int countA = 0;
                    int TG = 0;
                    int CP = 0;
                    int rowOther = 0;
                    var listPoint = db.sp_46_QCSelectWO_09_QCTAGSelect(QCNo1).ToList();
                    if (listPoint.Count > 0)
                    {
                        foreach (var rs in listPoint)
                        {
                            SOK = 0;
                            SNG = 0;
                            countA += 1;
                            TG = 0;
                            string[] PPTAG = rs.ofTAG.Split('o');
                            TG = Convert.ToInt32(PPTAG[0]);
                            // MessageBox.Show(countA.ToString());
                            if (TG > 0)
                            {
                                row1 = 6;
                                rowOther = 0;
                               // CP = TG;
                                if (TG <= 25)
                                {
                                    CP = TG;
                                }
                                else if (TG <= 50)
                                {
                                    CP = TG - 25;
                                }
                                else if (TG <= 75)
                                {
                                    CP = TG - 50;
                                }
                                else if (TG <= 100)
                                {
                                    CP = TG - 75;
                                }
                                else if (TG <= 125)
                                {
                                    CP = TG - 100;
                                }
                                else if (TG <= 150)
                                {
                                    CP = TG - 125;
                                }

                                var listPart = db.tb_QCGroupParts.Where(q => q.FormISO.Equals(FormISO) && q.PartNo.Equals(DValue.CODE)).OrderBy(o => o.Seq).ToList();
                                foreach (var rd in listPart)
                                {
                                    //Start Insert Checkmark  
                                    row1 += 1;

                                    
                                    

                                    //Start G=7,H=
                                    if (!rd.TopPic.Equals(""))
                                    {
                                        if(rd.TopPic.Equals("OTHER"))
                                        {
                                            rowOther = row1;
                                        }
                                        /////
                                        try
                                        {
                                            Remark = "";
                                            var gValue = db.sp_46_QCGetValue5601(rs.BarcodeTag, QHNo, rd.Seq).FirstOrDefault();
                                            PV = "P";

                                            if (gValue.CountA > 0)
                                            {

                                                PV = "O";
                                                if (gValue.CountA == 99)
                                                    PV = "";
                                            }
                                            var NValue = db.sp_46_QCGetValue55501(rs.BarcodeTag, QHNo, rd.Seq).FirstOrDefault();
                                            Remark = NValue.Remark;

                                            //Excel.Range Col0 = worksheet.get_Range(Getcolumn(CP+6) + row1.ToString(), Getcolumn(CP + 6) + row1.ToString());
                                            //Col0.Value2 = PV;
                                            if (TG <= 25)
                                            {
                                                if (PV.Equals("P"))
                                                {
                                                    if (rd.TopPic.Equals("OTHER"))
                                                    {
                                                        PV = "";
                                                    }
                                                }
                                                ////
                                                if (db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq) != "")
                                                {
                                                    PV = "O";
                                                    if (ValueInvalid == "")
                                                        ValueInvalid = db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);
                                                    else
                                                        ValueInvalid = ValueInvalid + "," + db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);

                                                }
                                                Excel.Range Col0 = worksheet.get_Range(Getcolumn(CP + 6) + row1.ToString(), Getcolumn(CP + 6) + row1.ToString());
                                                Col0.Value2 = PV;

                                            }
                                            else if (TG <= 50)
                                            {
                                                if (PV.Equals("P"))
                                                {
                                                    if (rd.TopPic.Equals("OTHER"))
                                                    {
                                                        PV = "";
                                                    }
                                                }
                                                ////
                                                if (db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq) != "")
                                                {
                                                    PV = "O";
                                                    if (ValueInvalid2 == "")
                                                        ValueInvalid2 = db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);
                                                    else
                                                        ValueInvalid2 = ValueInvalid + "," + db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);
                                                }

                                                Excel.Range Col02 = worksheet2.get_Range(Getcolumn(CP + 6) + row1.ToString(), Getcolumn(CP + 6) + row1.ToString());
                                                Col02.Value2 = PV;

                                            }
                                            else if (TG <= 75)
                                            {
                                                if (PV.Equals("P"))
                                                {
                                                    if (rd.TopPic.Equals("OTHER"))
                                                    {
                                                        PV = "";
                                                    }
                                                }
                                                ////
                                                if (db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq) != "")
                                                {
                                                    PV = "O";
                                                    if (ValueInvalid3 == "")
                                                        ValueInvalid3 = db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);
                                                    else
                                                        ValueInvalid3 = ValueInvalid + "," + db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);
                                                }

                                                Excel.Range Col03 = worksheet3.get_Range(Getcolumn(CP + 6) + row1.ToString(), Getcolumn(CP + 6) + row1.ToString());
                                                Col03.Value2 = PV;

                                            }
                                            else if (TG <= 100)
                                            {
                                                if (PV.Equals("P"))
                                                {
                                                    if (rd.TopPic.Equals("OTHER"))
                                                    {
                                                        PV = "";
                                                    }
                                                }
                                                ////
                                                if (db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq) != "")
                                                {
                                                    PV = "O";
                                                    if (ValueInvalid4 == "")
                                                        ValueInvalid4 = db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);
                                                    else
                                                        ValueInvalid4 = ValueInvalid + "," + db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);
                                                }

                                                Excel.Range Col04 = worksheet4.get_Range(Getcolumn(CP + 6) + row1.ToString(), Getcolumn(CP + 6) + row1.ToString());
                                                Col04.Value2 = PV;

                                            }
                                            else if (TG <= 125)
                                            {
                                                if (PV.Equals("P"))
                                                {
                                                    if (rd.TopPic.Equals("OTHER"))
                                                    {
                                                        PV = "";
                                                    }
                                                }
                                                ////
                                                if (db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq) != "")
                                                {
                                                    PV = "O";
                                                    if (ValueInvalid4 == "")
                                                        ValueInvalid4 = db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);
                                                    else
                                                        ValueInvalid4 = ValueInvalid + "," + db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);
                                                }

                                                Excel.Range Col05 = worksheet5.get_Range(Getcolumn(CP + 6) + row1.ToString(), Getcolumn(CP + 6) + row1.ToString());
                                                Col05.Value2 = PV;

                                            }
                                            else if (TG <= 150)
                                            {
                                                if (PV.Equals("P"))
                                                {
                                                    if (rd.TopPic.Equals("OTHER"))
                                                    {
                                                        PV = "";
                                                    }
                                                }
                                                ////
                                                if (db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq) != "")
                                                {
                                                    PV = "O";
                                                    if (ValueInvalid4 == "")
                                                        ValueInvalid4 = db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);
                                                    else
                                                        ValueInvalid4 = ValueInvalid + "," + db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);
                                                }

                                                Excel.Range Col06 = worksheet6.get_Range(Getcolumn(CP + 6) + row1.ToString(), Getcolumn(CP + 6) + row1.ToString());
                                                Col06.Value2 = PV;

                                            }

                                            if (!Remark.Equals(""))
                                            {
                                                if (TG <= 25)
                                                {
                                                    Excel.Range Col1 = worksheet.get_Range("AF" + Convert.ToString(row1));
                                                    Col1.Value2 = Remark;
                                                }
                                                else if (TG <= 50)
                                                {

                                                    Excel.Range Col12 = worksheet2.get_Range("AF" + Convert.ToString(row1));
                                                    Col12.Value2 = Remark;
                                                }
                                                else if (TG <= 75)
                                                {

                                                    Excel.Range Col13 = worksheet3.get_Range("AF" + Convert.ToString(row1));
                                                    Col13.Value2 = Remark;
                                                }
                                                else if (TG <= 100)
                                                {
                                                    Excel.Range Col14 = worksheet4.get_Range("AF" + Convert.ToString(row1));
                                                    Col14.Value2 = Remark;
                                                }
                                                else if (TG <= 125)
                                                {
                                                    Excel.Range Col15 = worksheet5.get_Range("AF" + Convert.ToString(row1));
                                                    Col15.Value2 = Remark;
                                                }
                                                else if (TG <= 150)
                                                {
                                                    Excel.Range Col16 = worksheet6.get_Range("AF" + Convert.ToString(row1));
                                                    Col16.Value2 = Remark;
                                                }
                                            }

                                        }
                                        catch (Exception ex) { MessageBox.Show(ex.Message); }
                                        //}




                                    }
                                    //SumNG//
                                }//foreach
                            }//cunt A

                            //Find count Tag
                            if (TG <= 25)
                            {
                                Excel.Range GNG = worksheet.get_Range(Getcolumn(CP + 6) + "22");
                                GNG.Value2 = db.get_QCSumQtyTAGNG(QHNo, rs.BarcodeTag, 3);
                                Excel.Range GOK = worksheet.get_Range(Getcolumn(CP + 6) + "21");
                                GOK.Value2 = db.get_QCSumQtyTAGNG(QHNo, rs.BarcodeTag, 4);
                                if (rowOther > 0)
                                {
                                    Excel.Range C15 = worksheet.get_Range("C" + rowOther.ToString());
                                    C15.Value2 = ValueInvalid;
                                }
                            }
                            else if (TG <= 50)
                            {
                                Excel.Range GNG2 = worksheet2.get_Range(Getcolumn(CP + 6) + "22");
                                GNG2.Value2 = db.get_QCSumQtyTAGNG(QHNo, rs.BarcodeTag, 3);
                                Excel.Range GOK2 = worksheet2.get_Range(Getcolumn(CP + 6) + "21");
                                GOK2.Value2 = db.get_QCSumQtyTAGNG(QHNo, rs.BarcodeTag, 4);
                                if (rowOther > 0)
                                {
                                    Excel.Range C15 = worksheet2.get_Range("C" + rowOther.ToString());
                                    C15.Value2 = ValueInvalid2;
                                }
                            }
                            else if (TG <= 75)
                            {
                                Excel.Range GNG3 = worksheet3.get_Range(Getcolumn(CP + 6) + "22");
                                GNG3.Value2 = db.get_QCSumQtyTAGNG(QHNo, rs.BarcodeTag, 3);
                                Excel.Range GOK3 = worksheet3.get_Range(Getcolumn(CP + 6) + "21");
                                GOK3.Value2 = db.get_QCSumQtyTAGNG(QHNo, rs.BarcodeTag, 4);
                                if (rowOther > 0)
                                {
                                    Excel.Range C15 = worksheet3.get_Range("C" + rowOther.ToString());
                                    C15.Value2 = ValueInvalid3;
                                }
                            }
                            else if (TG <= 100)
                            {
                                Excel.Range GNG4 = worksheet4.get_Range(Getcolumn(CP + 6) + "22");
                                GNG4.Value2 = db.get_QCSumQtyTAGNG(QHNo, rs.BarcodeTag, 3);
                                Excel.Range GOK4 = worksheet4.get_Range(Getcolumn(CP + 6) + "21");
                                GOK4.Value2 = db.get_QCSumQtyTAGNG(QHNo, rs.BarcodeTag, 4);
                                if (rowOther > 0)
                                {
                                    Excel.Range C15 = worksheet4.get_Range("C" + rowOther.ToString());
                                    C15.Value2 = ValueInvalid4;
                                }
                            }
                            else if (TG <= 125)
                            {
                                Excel.Range GNG5 = worksheet5.get_Range(Getcolumn(CP + 6) + "22");
                                GNG5.Value2 = db.get_QCSumQtyTAGNG(QHNo, rs.BarcodeTag, 3);
                                Excel.Range GOK5 = worksheet5.get_Range(Getcolumn(CP + 6) + "21");
                                GOK5.Value2 = db.get_QCSumQtyTAGNG(QHNo, rs.BarcodeTag, 4);
                                if (rowOther > 0)
                                {
                                    Excel.Range C15 = worksheet5.get_Range("C" + rowOther.ToString());
                                    C15.Value2 = ValueInvalid4;
                                }
                            }
                            else if (TG <= 150)
                            {
                                Excel.Range GNG6 = worksheet6.get_Range(Getcolumn(CP + 6) + "22");
                                GNG6.Value2 = db.get_QCSumQtyTAGNG(QHNo, rs.BarcodeTag, 3);
                                Excel.Range GOK6 = worksheet6.get_Range(Getcolumn(CP + 6) + "21");
                                GOK6.Value2 = db.get_QCSumQtyTAGNG(QHNo, rs.BarcodeTag, 4);
                                if (rowOther > 0)
                                {
                                    Excel.Range C15 = worksheet6.get_Range("C" + rowOther.ToString());
                                    C15.Value2 = ValueInvalid4;
                                }
                            }
                        }//for
                    }



                }

                excelBook.SaveAs(tempfile);
                excelBook.Close(false);
                excelApp.Quit();

                releaseObject(worksheet);
                releaseObject(worksheet2);
                releaseObject(worksheet3);
                releaseObject(worksheet4);
                releaseObject(worksheet5);
                releaseObject(worksheet6);
                releaseObject(excelBook);
                releaseObject(excelApp);
                Marshal.FinalReleaseComObject(worksheet);
                Marshal.FinalReleaseComObject(worksheet2);
                Marshal.FinalReleaseComObject(worksheet3);
                Marshal.FinalReleaseComObject(worksheet4);
                Marshal.FinalReleaseComObject(worksheet5);
                Marshal.FinalReleaseComObject(worksheet6);
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
            catch(Exception e) { MessageBox.Show(e.Message); }

        }
        public static void PrintData035(string WO, string PartNo, string QCNo1)
        {
            try
            {


                string DATA = AppDomain.CurrentDomain.BaseDirectory;
                string tempPath = System.IO.Path.GetTempPath();
                string FileName = "FM-PD-035.xlsx";
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
                Excel.Worksheet worksheet2 = (Excel.Worksheet)sheets.get_Item(2);
                Excel.Worksheet worksheet3 = (Excel.Worksheet)sheets.get_Item(3);
                Excel.Worksheet worksheet4 = (Excel.Worksheet)sheets.get_Item(4);

                // progressBar1.Maximum = 51;
                // progressBar1.Minimum = 1;
                int row1 = 6;               
                int Seq = 0;
                int TG = 0;           
                string PV = "P";
                string QHNo = QCNo1;
                string FormISO = "";
                //string cIssueBy1 = "";
               // string cIssueBy2 = "";
                string cCheckBy1 = "";
                string cCheckBy2 = "";
                string cCheckBy3 = "";
                string cCheckBy4 = "";
                string cCheckBy5 = "";
                string cCheckBy6 = "";
                string []SetData = new string[10];
                
                bool PAGE1 = true;
                bool PAGE2 = false;
                bool PAGE3 = false;
                bool PAGE4 = false;
                bool chek24 = true;
                string DN = "";
                string LotMark = "";// "Lot ที่ตอกสามารถอ่านได้อย่างชัดเจน ( " +")";
                string Line1Part = "";
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    //string Value1 = "";
                    //string Value2 = "";
                    //string LotNo = "";
                    string RefValue1 = "";
                    string RefValue2 = "";
                    string RefValue3 = "";
                    string PartName = "";
                   // string Remark = "";
                    string C9 = "";
                   // string ConnerElbo = "มุมการประกอบ Elbow กับ Cace อยู่ในค่าที่กำหนด";
                   
                    string GP5 = "";
                   
                    ///////////////SETValue/////////////////
                    var DValue = db.sp_46_QCSelectWO_01(WO).FirstOrDefault();
                    if (DValue != null)
                    {
                        var PTAGList = db.tb_QCTAGs.Where(p => p.QCNo.Equals(QHNo)).ToList();
                        if(PTAGList.Count>40)
                        {
                            PAGE2 = true;
                        }
                        if(PTAGList.Count>80)
                        {
                            PAGE3 = true;
                        }
                        if(PTAGList.Count>120)
                        {
                            PAGE4 = true;
                        }

                        if (PAGE1)
                        {
                            DN = DValue.DayNight;
                            PartName = DValue.NAME;
                            Excel.Range CStamp = worksheet.get_Range("Y3");
                            CStamp.Value2 = DValue.CODE;
                            Excel.Range CName = worksheet.get_Range("Y4");
                            CName.Value2 = DValue.NAME;

                            Excel.Range W5 = worksheet.get_Range("W5");
                            W5.Value2 = DValue.PORDER;

                            Excel.Range AE5 = worksheet.get_Range("AE5");
                            AE5.Value2 = DValue.LotNo;
                        }
                        if(PAGE2)
                        {
                            DN = DValue.DayNight;
                            PartName = DValue.NAME;
                            Excel.Range CStamp = worksheet2.get_Range("Y3");
                            CStamp.Value2 = DValue.CODE;
                            Excel.Range CName = worksheet2.get_Range("Y4");
                            CName.Value2 = DValue.NAME;

                            Excel.Range W5 = worksheet2.get_Range("W5");
                            W5.Value2 = DValue.PORDER;

                            Excel.Range AE5 = worksheet2.get_Range("AE5");
                            AE5.Value2 = DValue.LotNo;
                        }
                        if (PAGE3)
                        {
                            DN = DValue.DayNight;
                            PartName = DValue.NAME;

                            Excel.Range CStamp = worksheet3.get_Range("Y3");
                            CStamp.Value2 = DValue.CODE;

                            Excel.Range CName = worksheet3.get_Range("Y4");
                            CName.Value2 = DValue.NAME;

                            Excel.Range W5 = worksheet3.get_Range("W5");
                            W5.Value2 = DValue.PORDER;

                            Excel.Range AE5 = worksheet3.get_Range("AE5");
                            AE5.Value2 = DValue.LotNo;
                        }
                        if (PAGE4)
                        {
                            DN = DValue.DayNight;
                            PartName = DValue.NAME;

                            Excel.Range CStamp = worksheet4.get_Range("Y3");
                            CStamp.Value2 = DValue.CODE;

                            Excel.Range CName = worksheet4.get_Range("Y4");
                            CName.Value2 = DValue.NAME;

                            Excel.Range W5 = worksheet4.get_Range("W5");
                            W5.Value2 = DValue.PORDER;

                            Excel.Range AE5 = worksheet4.get_Range("AE5");
                            AE5.Value2 = DValue.LotNo;
                        }

                        LotMark = "Lot ที่ตอกสามารถอ่านได้อย่างชัดเจน (  "+ DValue.LotNo+"   )";
                        if (DValue.CODE.Length > 0)
                        {
                            if (dbClss.Right(DValue.CODE, 1).ToUpper().Equals("W"))
                            {
                                Line1Part = "Part No.ที่ Stamp ที่ CASE สามารถอ่านได้ชัดเจน  \n (   " + dbClss.Right(DValue.CODE, 8).Substring(0, 2) + " " + dbClss.Right(DValue.CODE, 6).Substring(0, 5) + "  )";
                            }
                            else
                            {
                                Line1Part = "Part No.ที่ Stamp ที่ CASE สามารถอ่านได้ชัดเจน  \n (   " + dbClss.Right(DValue.CODE, 7).Substring(0, 2) + " " + dbClss.Right(DValue.CODE, 5) + "  )";
                            }
                        }

                      

                        chek24 = true;
                        if (PartName.Contains("30-") || PartName.Contains("-30"))
                        {
                            chek24 = false;
                            GP5 = "30-24";
                        }
                        else
                        {
                            if (PartName.Contains("16-24"))
                            {
                                GP5 = "16-24";
                            }
                            else if (PartName.Contains("20-24"))
                            {
                                GP5 = "20-24";
                            }
                            else if (PartName.Contains("24-24"))
                            {
                                GP5 = "24-24";
                            }
                        }

                        





                        try
                        {
                            tb_QCHD qh = db.tb_QCHDs.Where(w => w.QCNo.Equals(QCNo1)).FirstOrDefault();
                            if (qh != null)
                            {

                                //////////Find UserName////////////
                                var uc = db.tb_QCCheckUsers.Where(u => u.QCNo.Equals(QCNo1)).ToList();                               
                                foreach (var rd in uc)
                                {
                                    DN = rd.DayN;
                                    if (DN.Equals("D"))
                                    {
                                        if (rd.UDesc.Equals("ผู้ตรวจสอบ"))
                                        {
                                            if (cCheckBy1.Equals(""))
                                                cCheckBy1 = rd.UserName;
                                            else
                                                cCheckBy1 = cCheckBy1 + "/" + rd.UserName;
                                            //DN1 = dbShowData.CheckDayN(Convert.ToDateTime(rd.ScanDate));
                                        }
                                        if (rd.UDesc.Equals("พนักงานตรวจ ก่อนผลิต"))
                                        {
                                            if (cCheckBy2.Equals(""))
                                                cCheckBy2 = rd.UserName;
                                            else
                                                cCheckBy2 = cCheckBy2 + "/" + rd.UserName;

                                           // DN2 = dbShowData.CheckDayN(Convert.ToDateTime(rd.ScanDate));
                                        }
                                        if (rd.UDesc.Equals("พนักงานตรวจ หลังผลิต"))
                                        {
                                            if (cCheckBy3.Equals(""))
                                                cCheckBy3 = rd.UserName;
                                            else
                                                cCheckBy3 = cCheckBy3 + "/" + rd.UserName;

                                           // DN3 = dbShowData.CheckDayN(Convert.ToDateTime(rd.ScanDate));
                                        }
                                    }else
                                    {
                                        if (rd.UDesc.Equals("ผู้ตรวจสอบ"))
                                        {

                                            if (cCheckBy4.Equals(""))
                                                cCheckBy4 = rd.UserName;
                                            else
                                                cCheckBy4 = cCheckBy4 + "/" + rd.UserName;
                                           // DN1 = dbShowData.CheckDayN(Convert.ToDateTime(rd.ScanDate));
                                        }
                                        if (rd.UDesc.Equals("พนักงานตรวจ ก่อนผลิต"))
                                        {
                                            if (cCheckBy5.Equals(""))
                                                cCheckBy5 = rd.UserName;
                                            else
                                                cCheckBy5 = cCheckBy5 + "/" + rd.UserName;

                                            //DN2 = dbShowData.CheckDayN(Convert.ToDateTime(rd.ScanDate));
                                        }
                                        if (rd.UDesc.Equals("พนักงานตรวจ หลังผลิต"))
                                        {
                                            if (cCheckBy6.Equals(""))
                                                cCheckBy6 = rd.UserName;
                                            else
                                                cCheckBy6 = cCheckBy6 + "/" + rd.UserName;

                                            //DN3 = dbShowData.CheckDayN(Convert.ToDateTime(rd.ScanDate));
                                        }
                                    }
                                }

                                FormISO = qh.FormISO;
                                QHNo = qh.QCNo;
                                RefValue1 = qh.RefValue1;
                                RefValue2 = qh.RefValue2;
                                RefValue3 = qh.RefValue3;

                                if (PAGE1)
                                {
                                    Excel.Range app = worksheet.get_Range("AJ4");
                                    app.Value2 = db.QC_GetUserName(qh.ApproveBy); //qh.ApproveBy;                                

                                    Excel.Range check1 = worksheet.get_Range("AT5");
                                    check1.Value2 = cCheckBy1;
                                    Excel.Range check4 = worksheet.get_Range("AW5");
                                    check4.Value2 = cCheckBy4;

                                    Excel.Range check2 = worksheet.get_Range("AO22");
                                    check2.Value2 = cCheckBy2;
                                    Excel.Range check5 = worksheet.get_Range("AT22");
                                    check5.Value2 = cCheckBy5;

                                    Excel.Range check3 = worksheet.get_Range("AO27");
                                    check3.Value2 = cCheckBy3;
                                    Excel.Range check6 = worksheet.get_Range("AT27");
                                    check6.Value2 = cCheckBy6;
                                    
                                    Excel.Range QD1 = worksheet.get_Range("K5");
                                    QD1.Value2 = Convert.ToDateTime(qh.CreateDate).ToString("dd") + " วัน " + Convert.ToDateTime(qh.CreateDate).ToString("MM") + " เดือน  " + Convert.ToDateTime(qh.CreateDate).ToString("yyyy") + " ปี";
                                    
                                    Excel.Range order = worksheet.get_Range("J4");
                                    order.Value2 = qh.OrderQty;// db.get_QCSumQtyTAGNG(qh.QCNo, "", 98);
                                    Excel.Range J16 = worksheet.get_Range("J16");
                                    J16.Value2 = GP5;

                                    Excel.Range KNG = worksheet.get_Range("K4");
                                    KNG.Value2 = db.get_QCSumQtyNG_RE(qh.QCNo, 1);
                                    Excel.Range Rework = worksheet.get_Range("M4");
                                    Rework.Value2 = db.get_QCSumQtyNG_RE(qh.QCNo, 2);
                                    

                                    Excel.Range B7 = worksheet.get_Range("B7");
                                    SetData[0]= db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 1);
                                    B7.Value2 = SetData[0];

                                    Excel.Range B8 = worksheet.get_Range("B8");
                                    B8.Value2 = Line1Part;// db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 2);

                                    Excel.Range B9 = worksheet.get_Range("B9");
                                    SetData[1]= db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 3);
                                    B9.Value2 = SetData[1];

                                    Excel.Range B10 = worksheet.get_Range("B10");
                                    SetData[2]= db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 4);
                                    B10.Value2 = SetData[2];

                                    Excel.Range B11 = worksheet.get_Range("B11");
                                    SetData[3]= db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 5) + " \n " + LotMark;
                                    B11.Value2 = SetData[3];

                                    Excel.Range B12 = worksheet.get_Range("B12");
                                    SetData[4]= db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 6);
                                    B12.Value2 = SetData[4];

                                    Excel.Range B13 = worksheet.get_Range("B13");
                                    SetData[5]= db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 7);
                                    B13.Value2 = SetData[5];

                                    Excel.Range B14 = worksheet.get_Range("B14");
                                    SetData[6]= db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 8);
                                    B14.Value2 = SetData[6];

                                    C9 = db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 9);
                                    Excel.Range B15 = worksheet.get_Range("B15");
                                    B15.Value2 = C9;
                                }

                                if(PAGE2)
                                {
                                    Excel.Range app = worksheet2.get_Range("AJ4");
                                    app.Value2 = db.QC_GetUserName(qh.ApproveBy); //qh.ApproveBy;                                

                                    Excel.Range check1 = worksheet2.get_Range("AT5");
                                    check1.Value2 = cCheckBy1;
                                    Excel.Range check4 = worksheet2.get_Range("AW5");
                                    check4.Value2 = cCheckBy4;

                                    Excel.Range check2 = worksheet2.get_Range("AO22");
                                    check2.Value2 = cCheckBy2;
                                    Excel.Range check5 = worksheet2.get_Range("AT22");
                                    check5.Value2 = cCheckBy5;

                                    Excel.Range check3 = worksheet2.get_Range("AO27");
                                    check3.Value2 = cCheckBy3;
                                    Excel.Range check6 = worksheet2.get_Range("AT27");
                                    check6.Value2 = cCheckBy6;

                                    Excel.Range QD1 = worksheet2.get_Range("K5");
                                    QD1.Value2 = Convert.ToDateTime(qh.CreateDate).ToString("dd") + " วัน " + Convert.ToDateTime(qh.CreateDate).ToString("MM") + " เดือน  " + Convert.ToDateTime(qh.CreateDate).ToString("yyyy") + " ปี";

                                    Excel.Range order = worksheet2.get_Range("J4");
                                    order.Value2 = qh.OrderQty;//db.get_QCSumQtyTAGNG(qh.QCNo, "", 99);
                                    Excel.Range J16 = worksheet2.get_Range("J16");
                                    J16.Value2 = GP5;

                                    Excel.Range KNG = worksheet2.get_Range("K4");
                                    KNG.Value2 = db.get_QCSumQtyNG_RE(qh.QCNo, 1);
                                    Excel.Range Rework = worksheet2.get_Range("M4");
                                    Rework.Value2 = db.get_QCSumQtyNG_RE(qh.QCNo, 2);

                                    Excel.Range B7 = worksheet2.get_Range("B7");
                                    B7.Value2 = SetData[0];// db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 1);

                                    Excel.Range B8 = worksheet2.get_Range("B8");
                                    B8.Value2 = Line1Part;// db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 2);

                                    Excel.Range B9 = worksheet2.get_Range("B9");
                                    B9.Value2 = SetData[1];// db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 3);

                                    Excel.Range B10 = worksheet2.get_Range("B10");
                                    B10.Value2 = SetData[2];// db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 4);

                                    Excel.Range B11 = worksheet2.get_Range("B11");
                                    B11.Value2 = SetData[3];// db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 5) + " \n " + LotMark;

                                    Excel.Range B12 = worksheet2.get_Range("B12");
                                    B12.Value2 = SetData[4];// db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 6);

                                    Excel.Range B13 = worksheet2.get_Range("B13");
                                    B13.Value2 = SetData[5];// db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 7);

                                    Excel.Range B14 = worksheet2.get_Range("B14");
                                    B14.Value2 = SetData[6];// db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 8);

                                  //  C9 = db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 9);
                                    Excel.Range B15 = worksheet2.get_Range("B15");
                                    B15.Value2 = C9;
                                }
                                if (PAGE3)
                                {
                                    Excel.Range app = worksheet3.get_Range("AJ4");
                                    app.Value2 = db.QC_GetUserName(qh.ApproveBy); //qh.ApproveBy;                                

                                    Excel.Range check1 = worksheet3.get_Range("AT5");
                                    check1.Value2 = cCheckBy1;
                                    Excel.Range check4 = worksheet3.get_Range("AW5");
                                    check4.Value2 = cCheckBy4;

                                    Excel.Range check2 = worksheet3.get_Range("AO22");
                                    check2.Value2 = cCheckBy2;
                                    Excel.Range check5 = worksheet3.get_Range("AT22");
                                    check5.Value2 = cCheckBy5;

                                    Excel.Range check3 = worksheet3.get_Range("AO27");
                                    check3.Value2 = cCheckBy3;
                                    Excel.Range check6 = worksheet3.get_Range("AT27");
                                    check6.Value2 = cCheckBy6;

                                    Excel.Range QD1 = worksheet3.get_Range("K5");
                                    QD1.Value2 = Convert.ToDateTime(qh.CreateDate).ToString("dd") + " วัน " + Convert.ToDateTime(qh.CreateDate).ToString("MM") + " เดือน  " + Convert.ToDateTime(qh.CreateDate).ToString("yyyy") + " ปี";

                                    Excel.Range order = worksheet3.get_Range("J4");
                                    order.Value2 = qh.OrderQty;//db.get_QCSumQtyTAGNG(qh.QCNo, "", 99);
                                    Excel.Range J16 = worksheet3.get_Range("J16");
                                    J16.Value2 = GP5;

                                    Excel.Range KNG = worksheet3.get_Range("K4");
                                    KNG.Value2 = db.get_QCSumQtyNG_RE(qh.QCNo, 1);
                                    Excel.Range Rework = worksheet3.get_Range("M4");
                                    Rework.Value2 = db.get_QCSumQtyNG_RE(qh.QCNo, 2);

                                    Excel.Range B7 = worksheet3.get_Range("B7");
                                    B7.Value2 = SetData[0];// db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 1);

                                    Excel.Range B8 = worksheet3.get_Range("B8");
                                    B8.Value2 = Line1Part;// db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 2);

                                    Excel.Range B9 = worksheet3.get_Range("B9");
                                    B9.Value2 = SetData[1];// db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 3);

                                    Excel.Range B10 = worksheet3.get_Range("B10");
                                    B10.Value2 = SetData[2];// db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 4);

                                    Excel.Range B11 = worksheet3.get_Range("B11");
                                    B11.Value2 = SetData[3];// db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 5) + " \n " + LotMark;

                                    Excel.Range B12 = worksheet3.get_Range("B12");
                                    B12.Value2 = SetData[4];// db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 6);

                                    Excel.Range B13 = worksheet3.get_Range("B13");
                                    B13.Value2 = SetData[5];// db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 7);

                                    Excel.Range B14 = worksheet3.get_Range("B14");
                                    B14.Value2 = SetData[6];// db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 8);

                                    //  C9 = db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 9);
                                    Excel.Range B15 = worksheet3.get_Range("B15");
                                    B15.Value2 = C9;
                                }
                                if(PAGE4)
                                {
                                    Excel.Range app = worksheet4.get_Range("AJ4");
                                    app.Value2 = db.QC_GetUserName(qh.ApproveBy); //qh.ApproveBy;                                

                                    Excel.Range check1 = worksheet4.get_Range("AT5");
                                    check1.Value2 = cCheckBy1;
                                    Excel.Range check4 = worksheet4.get_Range("AW5");
                                    check4.Value2 = cCheckBy4;

                                    Excel.Range check2 = worksheet4.get_Range("AO22");
                                    check2.Value2 = cCheckBy2;
                                    Excel.Range check5 = worksheet4.get_Range("AT22");
                                    check5.Value2 = cCheckBy5;

                                    Excel.Range check3 = worksheet4.get_Range("AO27");
                                    check3.Value2 = cCheckBy3;
                                    Excel.Range check6 = worksheet4.get_Range("AT27");
                                    check6.Value2 = cCheckBy6;

                                    Excel.Range QD1 = worksheet4.get_Range("K5");
                                    QD1.Value2 = Convert.ToDateTime(qh.CreateDate).ToString("dd") + " วัน " + Convert.ToDateTime(qh.CreateDate).ToString("MM") + " เดือน  " + Convert.ToDateTime(qh.CreateDate).ToString("yyyy") + " ปี";

                                    Excel.Range order = worksheet4.get_Range("J4");
                                    order.Value2 = qh.OrderQty;//db.get_QCSumQtyTAGNG(qh.QCNo, "", 99);
                                    Excel.Range J16 = worksheet4.get_Range("J16");
                                    J16.Value2 = GP5;

                                    Excel.Range KNG = worksheet4.get_Range("K4");
                                    KNG.Value2 = db.get_QCSumQtyNG_RE(qh.QCNo, 1);
                                    Excel.Range Rework = worksheet4.get_Range("M4");
                                    Rework.Value2 = db.get_QCSumQtyNG_RE(qh.QCNo, 2);

                                    Excel.Range B7 = worksheet4.get_Range("B7");
                                    B7.Value2 = SetData[0];// db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 1);

                                    Excel.Range B8 = worksheet4.get_Range("B8");
                                    B8.Value2 = Line1Part;// db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 2);

                                    Excel.Range B9 = worksheet4.get_Range("B9");
                                    B9.Value2 = SetData[1];// db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 3);

                                    Excel.Range B10 = worksheet4.get_Range("B10");
                                    B10.Value2 = SetData[2];// db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 4);

                                    Excel.Range B11 = worksheet4.get_Range("B11");
                                    B11.Value2 = SetData[3];// db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 5) + " \n " + LotMark;

                                    Excel.Range B12 = worksheet4.get_Range("B12");
                                    B12.Value2 = SetData[4];// db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 6);

                                    Excel.Range B13 = worksheet4.get_Range("B13");
                                    B13.Value2 = SetData[5];// db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 7);

                                    Excel.Range B14 = worksheet4.get_Range("B14");
                                    B14.Value2 = SetData[6];// db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 8);

                                    //  C9 = db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 9);
                                    Excel.Range B15 = worksheet4.get_Range("B15");
                                    B15.Value2 = C9;
                                }

                            }

                        }
                        catch (Exception ex) { MessageBox.Show("first " + ex.Message); }




                    }

                    ////////////////////////////////////////

                    int countA = 0;
                    int CountB = 0;
                    int CountC = 0;
                    int CountD = 0;
                    int TAG2 = 0;
                    int CA = 0;
                    int TG2 = 0;
                    int NGA = 0;
                    int NGB = 0;
                    int NGC = 0;
                    string TAGOf1 = "";
                    string TAGOf2 = "";
                    string TAGOf3 = "";

                    int CountTAG = 0;
                    var listPoint = db.sp_46_QCSelectWO_09_QCTAGSelect(QHNo).ToList();
                    CountTAG = listPoint.Count;
                    if (listPoint.Count > 0)
                    {
                        foreach (var rs in listPoint)
                        {
                            countA += 1;
                            if (countA > 40)
                            {
                                CountB += 1;
                            }
                            if(countA>80)
                            {
                                CountC += 1;
                            }
                            if (countA > 120)
                            {
                                CountD += 1;
                            }
                            

                            TG = 0;
                            
                            string[] PPTAG = rs.BarcodeTag.Split(',');
                            TG = Convert.ToInt32(PPTAG[2]);

                            //string[] PPTAG2 = rs.ofTAG.Split('o');
                            //TG2 = Convert.ToInt32(PPTAG2[0]);

                            if (chek24)
                            {
                                TAG2 += TG;
                            }
                            else
                            {
                                TAG2 += TG;
                            }
                            TG2 = 0;
                            TG2 = Convert.ToInt32(db.get_QCSumQtyTAGNG(QHNo, rs.BarcodeTag, 3));

                            if (listPoint.Count == countA)
                            {
                                NGA = TG;
                                TAGOf1 = PPTAG[5];
                            }
                            if ((listPoint.Count - 1) == countA)
                            {
                                NGB = TG;
                                TAGOf2 = PPTAG[5];
                            }
                            if ((listPoint.Count - 2) == countA)
                            {
                                NGC = TG;
                                TAGOf3 = PPTAG[5];
                            }

                            row1 = 6;
                                Seq = 0;
                                var listPart = db.tb_QCGroupParts.Where(q => q.FormISO.Equals(FormISO) && q.PartNo.Equals(DValue.CODE)).OrderBy(o => o.Seq).ToList();
                                CA = listPart.Count();
                                foreach (var rd in listPart)
                                {
                                    
                                    row1 += 1;
                                    Seq += 1;
                                if (!rd.SetData.Equals("") && row1 <= 15)
                                {
                                    try
                                    {

                                        var gValue = db.sp_46_QCGetValue5601(rs.BarcodeTag, QHNo, rd.Seq).FirstOrDefault();
                                        PV = "P";

                                        if (gValue.CountA > 0)
                                        {
                                            PV = "O";
                                            
                                            if (gValue.CountA == 99)
                                            {
                                                PV = "";
                                            }
                                        }
                                        if (rd.Seq.Equals(9) && C9.Equals(""))
                                        {
                                            PV = "";
                                        }

                                        if (countA <= 40)
                                        {

                                            Excel.Range Col0 = worksheet.get_Range(Getcolumn(countA + 10) + row1.ToString(), Getcolumn(countA + 10) + row1.ToString());
                                            Col0.Value2 = PV;
                                        }
                                        if (countA > 40 && countA <= 80)
                                        {
                                            Excel.Range Col0 = worksheet2.get_Range(Getcolumn(CountB + 10) + row1.ToString(), Getcolumn(CountB + 10) + row1.ToString());
                                            Col0.Value2 = PV;
                                        }
                                        if (countA > 80 && countA <= 120)
                                        {
                                            Excel.Range Col0 = worksheet3.get_Range(Getcolumn(CountC + 10) + row1.ToString(), Getcolumn(CountC + 10) + row1.ToString());
                                            Col0.Value2 = PV;
                                        }
                                        if (countA > 120 && countA <= 160)
                                        {
                                            Excel.Range Col0 = worksheet4.get_Range(Getcolumn(CountD + 10) + row1.ToString(), Getcolumn(CountD + 10) + row1.ToString());
                                            Col0.Value2 = PV;
                                        }





                                    }
                                    catch { }
                                     //catch (Exception ex) { MessageBox.Show("Mid " + ex.Message); }

                                }
                                //SumNG//       

                                if (countA <= 40)
                                {
                                    //NG Qty//

                                    ////////
                                    Excel.Range CSum = worksheet.get_Range(Getcolumn(countA + 10) + "17");
                                    CSum.Value2 = (TG - TG2);//.ToString();// TAG2.ToString();

                                    Excel.Range CSum1 = worksheet.get_Range(Getcolumn(countA + 10) + "18");
                                    CSum1.Value2 = TG2;//.ToString();// TAG2.ToString();

                                }
                                else if (countA > 40 && countA <= 80)
                                {
                                    if (PAGE2)
                                    {
                                        //NG Qty//

                                        ////////
                                        Excel.Range CSum = worksheet2.get_Range(Getcolumn(CountB + 10) + "17");
                                        CSum.Value2 = (TG - TG2);//.ToString();//TAG2.ToString(); ;

                                        Excel.Range CSum1 = worksheet2.get_Range(Getcolumn(CountB + 10) + "18");
                                        CSum1.Value2 = TG2;//.ToString();// TAG2.ToString();
                                    }
                                }
                                else if (countA > 80 && countA <= 120)
                                {
                                    if (PAGE3)
                                    {
                                        ////NG Qty//

                                        //////////
                                        Excel.Range CSum = worksheet3.get_Range(Getcolumn(CountC + 10) + "17");
                                        CSum.Value2 = (TG - TG2);//.ToString();// TAG2.ToString();

                                        Excel.Range CSum1 = worksheet3.get_Range(Getcolumn(CountC + 10) + "18");
                                        CSum1.Value2 = TG2;//.ToString();// TAG2.ToString();
                                    }
                                }
                                else if (countA > 120 && countA <= 160)
                                {
                                    if (PAGE4)
                                    {
                                        ////NG Qty//

                                        //////////
                                        Excel.Range CSum = worksheet4.get_Range(Getcolumn(CountD + 10) + "17");
                                        CSum.Value2 = (TG - TG2);//.ToString();// TAG2.ToString();

                                        Excel.Range CSum1 = worksheet4.get_Range(Getcolumn(CountD + 10) + "18");
                                        CSum1.Value2 = TG2;//.ToString();// TAG2.ToString();
                                    }
                                }
                                
                            }//foreach 
                            //}//cunt A //Page 1 End
                      
                        }//for

                        ////NGQty and Remark//
                        int RM = 0;
                        int TNG1 = 0;
                        int TNG2 = 0;
                        int TNG3 = 0;
                        tb_QCProblem qcp = db.tb_QCProblems.Where(p => p.QCNo.Equals(QHNo) && !p.NGQty.Equals(0)).FirstOrDefault();
                        if (qcp!=null)
                        {
                            var tgf = db.tb_QCTAGs.Where(s => s.QCNo.Equals(QHNo)).ToList();
                            foreach (var tf in tgf)
                            {
                                if (tf.ofTAG.Equals(TAGOf1))
                                {
                                    TNG1 = Convert.ToInt32(tf.NGQty);
                                }
                                else if (tf.ofTAG.Equals(TAGOf2))
                                {
                                    TNG2 = Convert.ToInt32(tf.NGQty);
                                }
                                else if (tf.ofTAG.Equals(TAGOf3))
                                {
                                    TNG3 = Convert.ToInt32(tf.NGQty);
                                }
                            }

                            if (countA <= 40)
                            {
                                //NG Qty//
                                //Excel.Range CSumA = worksheet.get_Range(Getcolumn(countA + 10) + "16");
                                //CSumA.Value2 = Convert.ToString(qcp.NGQty);
                                ////////
                                Excel.Range CSum = worksheet.get_Range("B16");
                                CSum.Value2 = qcp.ProblemName;
                                ///////////////////////////////
                                if (qcp.NGQty > 0)
                                {
                                    if(NGA<(qcp.NGQty+ TNG1))
                                    {
                                        Excel.Range CSum0 = worksheet.get_Range(Getcolumn(countA + 10) + "16");
                                        CSum0.Value2 = "O";
                                        Excel.Range CSumA = worksheet.get_Range(Getcolumn(countA + 10) + "17");
                                        CSumA.Value2 = 0;
                                        Excel.Range CSumB = worksheet.get_Range(Getcolumn(countA + 10) + "18");
                                        CSumB.Value2 = NGA;
                                        RM = (Convert.ToInt32(qcp.NGQty)+ TNG1) - NGA;
                                        if(RM>0)
                                        {
                                            if (NGB < (RM + TNG2))
                                            {
                                                Excel.Range CSum01 = worksheet.get_Range(Getcolumn(countA - 1 + 10) + "16");
                                                CSum01.Value2 = "O";
                                                Excel.Range CSumD = worksheet.get_Range(Getcolumn((countA - 1) + 10) + "17");
                                                CSumD.Value2 = 0;
                                                Excel.Range CSumE = worksheet.get_Range(Getcolumn((countA - 1) + 10) + "18");
                                                CSumE.Value2 = NGB;
                                                RM = (RM+ TNG2) - NGB;
                                                if (RM > 0)
                                                {
                                                    if (NGC < (RM+TNG3))
                                                    {
                                                        Excel.Range CSum02 = worksheet.get_Range(Getcolumn(countA - 2 + 10) + "16");
                                                        CSum02.Value2 = "O";
                                                        Excel.Range CSumF = worksheet.get_Range(Getcolumn((countA - 2) + 10) + "17");
                                                        CSumF.Value2 = 0;
                                                        Excel.Range CSumG = worksheet.get_Range(Getcolumn((countA - 2) + 10) + "18");
                                                        CSumG.Value2 = NGC;
                                                    }
                                                    else
                                                    {
                                                        Excel.Range CSum02 = worksheet.get_Range(Getcolumn(countA - 2 + 10) + "16");
                                                        CSum02.Value2 = "O";
                                                        Excel.Range CSumF = worksheet.get_Range(Getcolumn((countA - 2) + 10) + "17");
                                                        CSumF.Value2 = NGC - (RM+ TNG3);
                                                        Excel.Range CSumG = worksheet.get_Range(Getcolumn((countA - 2) + 10) + "18");
                                                        CSumG.Value2 = RM+ TNG3;
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                Excel.Range CSum01 = worksheet.get_Range(Getcolumn(countA - 1 + 10) + "16");
                                                CSum01.Value2 = "O";
                                                Excel.Range CSumD = worksheet.get_Range(Getcolumn((countA - 1) + 10) + "17");
                                                CSumD.Value2 = NGB - (RM + TNG2);
                                                Excel.Range CSumE = worksheet.get_Range(Getcolumn((countA - 1) + 10) + "18");
                                                CSumE.Value2 = (RM + TNG2);
                                            }
                                        }
                                                                                

                                    }
                                    else
                                    {
                                        Excel.Range CSum0 = worksheet.get_Range(Getcolumn(countA + 10) + "16");
                                        CSum0.Value2 = "O";
                                        Excel.Range CSumA = worksheet.get_Range(Getcolumn(countA + 10) + "17");
                                        CSumA.Value2 = NGA - (qcp.NGQty+ TNG1);
                                        Excel.Range CSumB = worksheet.get_Range(Getcolumn(countA + 10) + "18");
                                        CSumB.Value2 = (qcp.NGQty+ TNG1);
                                    }
                                    
                                }

                            }
                            else if (countA > 40 && countA <= 80)
                            {
                                if (PAGE2)
                                {
                                    //NG Qty//
                                    //Excel.Range CSumA = worksheet2.get_Range(Getcolumn(CountB + 10) + "16");
                                    //CSumA.Value2 = Convert.ToString(qcp.NGQty);
                                    ////////
                                    Excel.Range CSum = worksheet2.get_Range("B16");
                                    CSum.Value2 = qcp.ProblemName;

                                    if (qcp.NGQty > 0)
                                    {
                                        if (NGA < (qcp.NGQty+ TNG1))
                                        {
                                            Excel.Range CSum0 = worksheet2.get_Range(Getcolumn(CountB + 10) + "16");
                                            CSum0.Value2 = "O";
                                            Excel.Range CSumA = worksheet2.get_Range(Getcolumn(CountB + 10) + "17");
                                            CSumA.Value2 = 0;
                                            Excel.Range CSumB = worksheet2.get_Range(Getcolumn(CountB + 10) + "18");
                                            CSumB.Value2 = NGA;
                                            RM = (Convert.ToInt32(qcp.NGQty)+ TNG1) - NGA;
                                            if (RM > 0)
                                            {
                                                if (NGB < (RM+ TNG2))
                                                {
                                                    Excel.Range CSum1 = worksheet2.get_Range(Getcolumn((CountB - 1) + 10) + "16");
                                                    CSum1.Value2 = "O";
                                                    Excel.Range CSumD = worksheet2.get_Range(Getcolumn((CountB - 1) + 10) + "17");
                                                    CSumD.Value2 = 0;
                                                    Excel.Range CSumE = worksheet2.get_Range(Getcolumn((CountB - 1) + 10) + "18");
                                                    CSumE.Value2 = NGB;
                                                    RM = (RM+ TNG2) - NGB;
                                                    if (RM > 0)
                                                    {
                                                        if (NGC < (RM+ TNG3))
                                                        {
                                                            Excel.Range CSum2 = worksheet2.get_Range(Getcolumn((CountB - 2) + 10) + "16");
                                                            CSum2.Value2 = "O";
                                                            Excel.Range CSumF = worksheet2.get_Range(Getcolumn((CountB - 2) + 10) + "17");
                                                            CSumF.Value2 = 0;
                                                            Excel.Range CSumG = worksheet2.get_Range(Getcolumn((CountB - 2) + 10) + "18");
                                                            CSumG.Value2 = NGC;
                                                        }
                                                        else
                                                        {
                                                            Excel.Range CSum2 = worksheet2.get_Range(Getcolumn((CountB - 2) + 10) + "16");
                                                            CSum2.Value2 = "O";
                                                            Excel.Range CSumF = worksheet2.get_Range(Getcolumn((CountB - 2) + 10) + "17");
                                                            CSumF.Value2 = NGC - (RM + TNG3);
                                                            Excel.Range CSumG = worksheet2.get_Range(Getcolumn((CountB - 2) + 10) + "18");
                                                            CSumG.Value2 = RM+ TNG3;
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    Excel.Range CSum1 = worksheet2.get_Range(Getcolumn((CountB-1) + 10) + "16");
                                                    CSum1.Value2 = "O";
                                                    Excel.Range CSumD = worksheet2.get_Range(Getcolumn((CountB - 1) + 10) + "17");
                                                    CSumD.Value2 = NGB - (RM+ TNG2);
                                                    Excel.Range CSumE = worksheet2.get_Range(Getcolumn((CountB - 1) + 10) + "18");
                                                    CSumE.Value2 = RM+ TNG2;
                                                }
                                            }


                                        }
                                        else
                                        {
                                            Excel.Range CSum0 = worksheet2.get_Range(Getcolumn(CountB + 10) + "16");
                                            CSum0.Value2 = "O";
                                            Excel.Range CSumA = worksheet2.get_Range(Getcolumn(CountB + 10) + "17");
                                            CSumA.Value2 = NGA - (qcp.NGQty+ TNG1);
                                            Excel.Range CSumB = worksheet2.get_Range(Getcolumn(CountB + 10) + "18");
                                            CSumB.Value2 = qcp.NGQty+ TNG1;
                                        }

                                    }
                                }
                            }
                            else if (countA > 80 && countA <= 120)
                            {
                                if (PAGE3)
                                {
                                    //NG Qty//
                                    //Excel.Range CSumA = worksheet3.get_Range(Getcolumn(CountC + 10) + "16");
                                    //CSumA.Value2 = Convert.ToString(qcp.NGQty);
                                    ////////
                                    Excel.Range CSum = worksheet3.get_Range("B16");
                                    CSum.Value2 = qcp.ProblemName;

                                    if (qcp.NGQty > 0)
                                    {
                                        if (NGA < (qcp.NGQty+ TNG1))
                                        {
                                            Excel.Range CSum0 = worksheet3.get_Range(Getcolumn(CountC + 10) + "16");
                                            CSum0.Value2 = "O";
                                            Excel.Range CSumA = worksheet3.get_Range(Getcolumn(CountC + 10) + "17");
                                            CSumA.Value2 = 0;
                                            Excel.Range CSumB = worksheet3.get_Range(Getcolumn(CountC + 10) + "18");
                                            CSumB.Value2 = NGA;
                                            RM = (Convert.ToInt32(qcp.NGQty)+ TNG1) - NGA;
                                            if (RM > 0)
                                            {
                                                if (NGB < (RM+ TNG2))
                                                {
                                                    Excel.Range CSum1 = worksheet3.get_Range(Getcolumn((CountC - 1) + 10) + "16");
                                                    CSum1.Value2 = "O";
                                                    Excel.Range CSumD = worksheet3.get_Range(Getcolumn((CountC - 1) + 10) + "17");
                                                    CSumD.Value2 = 0;
                                                    Excel.Range CSumE = worksheet3.get_Range(Getcolumn((CountC - 1) + 10) + "18");
                                                    CSumE.Value2 = NGB;
                                                    RM = (RM+ TNG2) - NGB;
                                                    if (RM > 0)
                                                    {
                                                        if (NGC < (RM+ TNG3))
                                                        {
                                                            Excel.Range CSum2 = worksheet3.get_Range(Getcolumn((CountC - 2) + 10) + "16");
                                                            CSum2.Value2 = "O";
                                                            Excel.Range CSumF = worksheet3.get_Range(Getcolumn((CountC - 2) + 10) + "17");
                                                            CSumF.Value2 = 0;
                                                            Excel.Range CSumG = worksheet3.get_Range(Getcolumn((CountC - 2) + 10) + "18");
                                                            CSumG.Value2 = NGC;
                                                        }
                                                        else
                                                        {
                                                            Excel.Range CSum2 = worksheet3.get_Range(Getcolumn((CountC - 2) + 10) + "16");
                                                            CSum2.Value2 = "O";
                                                            Excel.Range CSumF = worksheet3.get_Range(Getcolumn((CountC - 2) + 10) + "17");
                                                            CSumF.Value2 = NGC - (RM + TNG3);
                                                            Excel.Range CSumG = worksheet3.get_Range(Getcolumn((CountC - 2) + 10) + "18");
                                                            CSumG.Value2 = RM+ TNG3;
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    Excel.Range CSum1 = worksheet3.get_Range(Getcolumn((CountC-1) + 10) + "16");
                                                    CSum1.Value2 = "O";
                                                    Excel.Range CSumD = worksheet3.get_Range(Getcolumn((CountC - 1) + 10) + "17");
                                                    CSumD.Value2 = NGB - (RM+ TNG2);
                                                    Excel.Range CSumE = worksheet3.get_Range(Getcolumn((CountC - 1) + 10) + "18");
                                                    CSumE.Value2 = RM+ TNG2;
                                                }
                                            }


                                        }
                                        else
                                        {
                                            Excel.Range CSum0 = worksheet3.get_Range(Getcolumn(CountC + 10) + "16");
                                            CSum0.Value2 = "O";
                                            Excel.Range CSumA = worksheet3.get_Range(Getcolumn(CountC + 10) + "17");
                                            CSumA.Value2 = NGA - (qcp.NGQty+ TNG1);
                                            Excel.Range CSumB = worksheet3.get_Range(Getcolumn(CountC + 10) + "18");
                                            CSumB.Value2 = qcp.NGQty+ TNG1;
                                        }

                                    }
                                }
                            }
                            else if (countA > 120 && countA <= 160)
                            {
                                if (PAGE4)
                                {
                                    //NG Qty//
                                    //Excel.Range CSumA = worksheet3.get_Range(Getcolumn(CountC + 10) + "16");
                                    //CSumA.Value2 = Convert.ToString(qcp.NGQty);
                                    ////////
                                    Excel.Range CSum = worksheet4.get_Range("B16");
                                    CSum.Value2 = qcp.ProblemName;

                                    if (qcp.NGQty > 0)
                                    {
                                        if (NGA < (qcp.NGQty + TNG1))
                                        {
                                            Excel.Range CSum0 = worksheet4.get_Range(Getcolumn(CountD + 10) + "16");
                                            CSum0.Value2 = "O";
                                            Excel.Range CSumA = worksheet4.get_Range(Getcolumn(CountD + 10) + "17");
                                            CSumA.Value2 = 0;
                                            Excel.Range CSumB = worksheet4.get_Range(Getcolumn(CountD + 10) + "18");
                                            CSumB.Value2 = NGA;
                                            RM = (Convert.ToInt32(qcp.NGQty) + TNG1) - NGA;
                                            if (RM > 0)
                                            {
                                                if (NGB < (RM + TNG2))
                                                {
                                                    Excel.Range CSum1 = worksheet4.get_Range(Getcolumn((CountD - 1) + 10) + "16");
                                                    CSum1.Value2 = "O";
                                                    Excel.Range CSumD = worksheet4.get_Range(Getcolumn((CountD - 1) + 10) + "17");
                                                    CSumD.Value2 = 0;
                                                    Excel.Range CSumE = worksheet4.get_Range(Getcolumn((CountD - 1) + 10) + "18");
                                                    CSumE.Value2 = NGB;
                                                    RM = (RM + TNG2) - NGB;
                                                    if (RM > 0)
                                                    {
                                                        if (NGC < (RM + TNG3))
                                                        {
                                                            Excel.Range CSum2 = worksheet4.get_Range(Getcolumn((CountD - 2) + 10) + "16");
                                                            CSum2.Value2 = "O";
                                                            Excel.Range CSumF = worksheet4.get_Range(Getcolumn((CountD - 2) + 10) + "17");
                                                            CSumF.Value2 = 0;
                                                            Excel.Range CSumG = worksheet4.get_Range(Getcolumn((CountD - 2) + 10) + "18");
                                                            CSumG.Value2 = NGC;
                                                        }
                                                        else
                                                        {
                                                            Excel.Range CSum2 = worksheet4.get_Range(Getcolumn((CountD - 2) + 10) + "16");
                                                            CSum2.Value2 = "O";
                                                            Excel.Range CSumF = worksheet4.get_Range(Getcolumn((CountD - 2) + 10) + "17");
                                                            CSumF.Value2 = NGC - (RM + TNG3);
                                                            Excel.Range CSumG = worksheet4.get_Range(Getcolumn((CountD - 2) + 10) + "18");
                                                            CSumG.Value2 = RM + TNG3;
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    Excel.Range CSum1 = worksheet4.get_Range(Getcolumn((CountD - 1) + 10) + "16");
                                                    CSum1.Value2 = "O";
                                                    Excel.Range CSumD = worksheet4.get_Range(Getcolumn((CountD - 1) + 10) + "17");
                                                    CSumD.Value2 = NGB - (RM + TNG2);
                                                    Excel.Range CSumE = worksheet4.get_Range(Getcolumn((CountD - 1) + 10) + "18");
                                                    CSumE.Value2 = RM + TNG2;
                                                }
                                            }


                                        }
                                        else
                                        {
                                            Excel.Range CSum0 = worksheet4.get_Range(Getcolumn(CountD + 10) + "16");
                                            CSum0.Value2 = "O";
                                            Excel.Range CSumA = worksheet4.get_Range(Getcolumn(CountD + 10) + "17");
                                            CSumA.Value2 = NGA - (qcp.NGQty + TNG1);
                                            Excel.Range CSumB = worksheet4.get_Range(Getcolumn(CountD + 10) + "18");
                                            CSumB.Value2 = qcp.NGQty + TNG1;
                                        }

                                    }
                                }
                            }
                        }

                        ////// PC Check ///
                    }



                }

                excelBook.SaveAs(tempfile);
                excelBook.Close(false);
                excelApp.Quit();
                releaseObject(worksheet);
                releaseObject(worksheet2);
                releaseObject(worksheet3);
                releaseObject(excelBook);
                releaseObject(excelApp);

                Marshal.FinalReleaseComObject(worksheet);                
                Marshal.FinalReleaseComObject(worksheet2);
                Marshal.FinalReleaseComObject(worksheet3);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet2);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet3);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                GC.GetTotalMemory(false);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.GetTotalMemory(true);
                System.Diagnostics.Process.Start(tempfile);

            }
            catch(Exception ex) { MessageBox.Show("last "+ex.Message); }

        }
        public static void PrintData033(string WO, string PartNo, string QCNo1)
        {
            try
            {
                string DATA = AppDomain.CurrentDomain.BaseDirectory;
                string tempPath = System.IO.Path.GetTempPath();
                string FileName = "FM-PD-033.xls";
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
                int row1 = 14;
               
                string PV = "P";
                string QHNo = QCNo1;
                string FormISO = "";
               
                string DN = "";
                string cIssueBy1 = "";
                string cIssueBy2 = "";
                string cCheckBy1 = "";
                string cCheckBy2 = "";
                string cCheckBy3 = "";
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                   
                    string LotNo = "";
                    
                    string RefValue2 = "";
                    string RefValue3 = "";
                    string PartName = "";
                   
                    bool chek24 = true;
                    string HeaderValue1 = "Piggy Back Checksheet การตรวจสอบด้วยตนเอง　（Size 24）";
                    string HeaderValue2 = "Piggy Back Checksheet การตรวจสอบด้วยตนเอง　（Size 30）";
                    ///////////////SETValue/////////////////
                    var DValue = db.sp_46_QCSelectWO_01(WO).FirstOrDefault();
                    if (DValue != null)
                    {
                        DN = DValue.DayNight;
                        PartName = DValue.NAME;
                        Excel.Range CStamp = worksheet.get_Range("O6");
                        CStamp.Value2 = DValue.CODE;

                        Excel.Range WWo = worksheet.get_Range("U5");
                        WWo.Value2 = DValue.PORDER;

                        //Excel.Range CName = worksheet.get_Range("C3");
                        //CName.Value2 = DValue.DeliveryDate;
                        //Excel.Range QD = worksheet.get_Range("D3");
                        //QD.Value2 = DValue.OrderQty;
                        //Excel.Range CDate = worksheet.get_Range("I3");
                        //CDate.Value2 = DValue.LotNo;
                        chek24 = true;
                        RefValue2 = "6300～7300　N";
                        RefValue3 = "5320～6200　N";
                        if (PartName.Contains("-30") || PartName.Contains("30-"))
                        {
                            chek24 = false;
                            RefValue2 = "８２００～９７２０　N";
                            RefValue3 = "７０６０～８４２０　N";
                        }

                        Excel.Range Header1 = worksheet.get_Range("I1");
                        Excel.Range Line19 = worksheet.get_Range("I19");
                        Excel.Range LIne20 = worksheet.get_Range("I20");
                        if (chek24)
                        {
                            Header1.Value2 = HeaderValue1;
                            Line19.Value2 = RefValue2;
                            LIne20.Value2 = RefValue3;
                        }
                        else
                        {
                            Header1.Value2 = HeaderValue2;
                            Line19.Value2 = RefValue2;
                            LIne20.Value2 = RefValue3;
                        }




                        try
                        {
                            tb_QCHD qh = db.tb_QCHDs.Where(w => w.QCNo.Equals(QCNo1)).FirstOrDefault();
                            if (qh != null)
                            {
                                FormISO = qh.FormISO;
                                LotNo = qh.LotNo;

                                //////////Find UserName////////////
                                var uc = db.tb_QCCheckUsers.Where(u => u.QCNo.Equals(QCNo1)).ToList();
                                int r2 = 0;
                                int r3 = 0;
                                foreach (var rd in uc)
                                {
                                    DN = dbShowData.CheckDayN(Convert.ToDateTime(rd.ScanDate));                                   

                                    if (rd.UDesc.Equals("ผู้ตรวจสอบ หัว"))
                                        cCheckBy1 = rd.UserName;
                                    if (rd.UDesc.Equals("ผู้ตรวจสอบ กลาง"))
                                        cCheckBy2 = rd.UserName;
                                    if (rd.UDesc.Equals("ผู้ตรวจสอบ ท้าย"))
                                        cCheckBy3 = rd.UserName;


                                }

                                //Excel.Range CDate1 = worksheet.get_Range("AH2");
                                //CDate1.Value2 = Convert.ToDateTime(qh.CreateDate).ToString("yyyy");
                                //Excel.Range CDate2 = worksheet.get_Range("AK2");
                                //CDate2.Value2 = Convert.ToDateTime(qh.CreateDate).ToString("MM");
                                //Excel.Range CDate3 = worksheet.get_Range("AN2");
                                //CDate3.Value2 = Convert.ToDateTime(qh.CreateDate).ToString("dd");

                                Excel.Range AD2 = worksheet.get_Range("AD2");
                                AD2.Value2 = Convert.ToDateTime(qh.CreateDate).ToString("yyyy")+"　　ปี     "+ Convert.ToDateTime(qh.CreateDate).ToString("MM") + "　 เดือน     "+ Convert.ToDateTime(qh.CreateDate).ToString("dd") + "　　วัน";
                                
                                Excel.Range Ap = worksheet.get_Range("AP4");
                                Ap.Value2 = db.QC_GetUserName(qh.ApproveBy);// qh.ApproveBy;

                                Excel.Range O9 = worksheet.get_Range("O9");
                                O9.Value2 = cCheckBy1;

                                Excel.Range Q9 = worksheet.get_Range("Q9");
                                Q9.Value2 = cCheckBy2;

                                Excel.Range S9 = worksheet.get_Range("S9");
                                S9.Value2 = cCheckBy3;
                                
                                QHNo = qh.QCNo;
                                // RefValue1 = qh.RefValue1;
                                //   RefValue2 = qh.RefValue2;
                                //  RefValue3 = qh.RefValue3;
                                

                                Excel.Range O19 = worksheet.get_Range("O19");
                                O19.Value2 = db.get_QC_DATAPoint(qh.QCNo, "", 8);

                                Excel.Range O20 = worksheet.get_Range("O20");
                                O20.Value2 = db.get_QC_DATAPoint(qh.QCNo, "", 9);
                            }

                        }
                        catch { }

                    }

                    ////////////////////////////////////////
                    

                    int countA = 0;
                    string Colm = "";
                    string ValuePoint = "";
                    var listPoint = db.sp_46_QCSelectWO_09_QCTAGSelect(QHNo).ToList();
                    if (listPoint.Count > 0)
                    {
                        foreach (var rs in listPoint)
                        {
                            countA += 1;
                            // MessageBox.Show(countA.ToString());
                            if (rs.Seq == 1)
                            {

                                ValuePoint=db.get_QC_DATAPoint(rs.QCNo, rs.BarcodeTag, 1);
                            }
                            if (countA <= 3)
                            {
                                row1 = 11;
                                var listPart = db.tb_QCGroupParts.Where(q => q.FormISO.Equals(FormISO)).OrderBy(o => o.Seq).ToList();
                                foreach (var rd in listPart)
                                {
                                    //Start Insert Checkmark  
                                    row1 += 1;
                                    //Start G=7,H=
                                    if (!rd.SetData.Equals("") && row1 <= 21)
                                    {
                                        try
                                        {
                                            
                                            var gValue = db.sp_46_QCGetValue5601(rs.BarcodeTag, QHNo, rd.Seq).FirstOrDefault();
                                            PV = "P";

                                            if (gValue.CountA > 0)
                                            {
                                                PV = "O";
                                                if (gValue.CountA == 99)
                                                    PV = "";
                                            }
                                            if(rd.Seq.Equals(6))
                                            {
                                                PV = "";
                                            }

                                          //  if (row1 == 19)
                                          //      row1 += 5;

                                            if (countA == 1)
                                                Colm = "O";

                                            else if (countA == 2)
                                                Colm = "Q";
                                            else
                                                Colm = "S";
                                            if (row1 == 19 || row1 == 20 || row1==17)
                                            {

                                            }
                                            else if (row1 == 12)
                                            {
                                                Excel.Range Col0 = worksheet.get_Range(Colm + "" + row1.ToString());
                                                Col0.Value2 = ValuePoint;
                                            }
                                            else if (row1 == 15)
                                            {
                                                if (gValue.CountA < 99)
                                                {
                                                    Excel.Range Col0 = worksheet.get_Range(Colm + "" + row1.ToString());
                                                    Col0.Value2 = LotNo;
                                                }
                                            }
                                            else
                                            {
                                                if (row1 != 9 && row1 != 8)
                                                {
                                                    Excel.Range Col0 = worksheet.get_Range(Colm + "" + row1.ToString());
                                                    Col0.Value2 = PV;
                                                }
                                            }

                                            




                                        }
                                        catch (Exception ex) { MessageBox.Show(ex.Message); }

                                    }
                                    //SumNG//

                                    //  NGQ = db.get_QCSumQtyTAGNG(QHNo,rs.BarcodeTag,
                                    //Excel.Range CSum = worksheet.get_Range(Getcolumn(countA + 6) + "26");
                                    //CSum.Value2 = rs.NGQty;



                                }//foreach
                            }//cunt A
                        }//for
                    }



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

        }
        public static void PrintFMPD112(string WO, string PartNo, string QCNo1, string FromIS)
        {
            try
            {
                string DATA = AppDomain.CurrentDomain.BaseDirectory;
                string tempPath = System.IO.Path.GetTempPath();
                string FileName = "FM-PD-112.xlsx";
                string TypeReport = GetReportName("HINO", PartNo, FromIS).ToUpper();
                string HeaderText = "CHECK SHEET ตรวจสอบ100% " + Environment.NewLine + " ";
                if (TypeReport.Equals("HINO"))
                {
                    HeaderText += "S-CAM PIGGY BACK  (HINO)";
                }
                else
                {
                    HeaderText += "S-CAM PIGGY BACK  (ISUZU)";
                }
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
                Excel.Worksheet worksheet2 = (Excel.Worksheet)sheets.get_Item(2);
                Excel.Worksheet worksheet3 = (Excel.Worksheet)sheets.get_Item(3);
                Excel.Worksheet worksheet4 = (Excel.Worksheet)sheets.get_Item(4);
                // progressBar1.Maximum = 51;
                // progressBar1.Minimum = 1;
                int row1 = 6;
                int Seq = 0;
                int TG = 0;
                string PV = "P";
                string QHNo = QCNo1;
                string FormISO = "";
                string[] Inspector = new string[10];
                string cCheckBy1 = "";
                string cCheckBy2 = "";
                string cCheckBy3 = "";
                string cCheckBy4 = "";
                string cCheckBy5 = "";
                string cCheckBy6 = "";
                string[] SetData = new string[10];
                bool PAGE1 = true;
                bool PAGE2 = false;
                bool PAGE3 = false;
                bool PAGE4 = false;
              //  bool chek24 = true;

                string DN = "";
                string LotMark = "LOT ต้องถูกต้อง, มีความชัดเจนสามารถอ่านได้ " + Environment.NewLine + " " ;
                string Line1Part = "";
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    string RefValue1 = "";
                    string RefValue2 = "";
                    string RefValue3 = "";
                    string PartName = "";
                    string C9 = "";
                    string GP5 = "";
                    ///////////////SETValue/////////////////////////////////
                    var DValue = db.sp_46_QCSelectWO_01(WO).FirstOrDefault();
                    if (DValue != null)
                    {
                        LotMark = LotMark + " ( " + DValue.LotNo + " )";
                        var PTAGList = db.tb_QCTAGs.Where(p => p.QCNo.Equals(QHNo)).ToList();
                        if (PTAGList.Count > 40)
                        {
                            PAGE2 = true;
                        }
                        if (PTAGList.Count > 80)
                        {
                            PAGE3 = true;
                        }
                        if (PTAGList.Count > 120)
                        {
                            PAGE4 = true;
                        }

                        DN = DValue.DayNight;
                        PartName = DValue.NAME;
                        if (PAGE1)
                        {
                            InsertToExcel(ref worksheet, "A4", HeaderText);
                            InsertToExcel(ref worksheet, "Y3", DValue.CODE);
                            InsertToExcel(ref worksheet, "Y4", DValue.NAME);
                            InsertToExcel(ref worksheet, "W5", DValue.PORDER);
                            InsertToExcel(ref worksheet, "AE5", DValue.LotNo);
                        }
                        if (PAGE2)
                        {
                            InsertToExcel(ref worksheet2, "A4", HeaderText);
                            InsertToExcel(ref worksheet2, "Y3", DValue.CODE);
                            InsertToExcel(ref worksheet2, "Y4", DValue.NAME);
                            InsertToExcel(ref worksheet2, "W5", DValue.PORDER);
                            InsertToExcel(ref worksheet2, "AE5", DValue.LotNo);
                        }
                        if (PAGE3)
                        {
                            InsertToExcel(ref worksheet3, "A4", HeaderText);
                            InsertToExcel(ref worksheet3, "Y3", DValue.CODE);
                            InsertToExcel(ref worksheet3, "Y4", DValue.NAME);
                            InsertToExcel(ref worksheet3, "W5", DValue.PORDER);
                            InsertToExcel(ref worksheet3, "AE5", DValue.LotNo);
                        }
                        if (PAGE4)
                        {
                            InsertToExcel(ref worksheet4, "A4", HeaderText);
                            InsertToExcel(ref worksheet4, "Y3", DValue.CODE);
                            InsertToExcel(ref worksheet4, "Y4", DValue.NAME);
                            InsertToExcel(ref worksheet4, "W5", DValue.PORDER);
                            InsertToExcel(ref worksheet4, "AE5", DValue.LotNo);
                        }


                        try
                        {
                            tb_QCHD qh = db.tb_QCHDs.Where(w => w.QCNo.Equals(QCNo1)).FirstOrDefault();
                            if (qh != null)
                            {
                                //////////Find UserName////////////
                                var uc = db.tb_QCCheckUsers.Where(u => u.QCNo.Equals(QCNo1)).ToList();
                                foreach (var rd in uc)
                                {
                                    DN = rd.DayN;
                                    if (DN.Equals("D"))
                                    {
                                        if (rd.UDesc.Equals("ผู้ตรวจสอบ"))
                                        {
                                            if (cCheckBy1.Equals(""))
                                                cCheckBy1 = rd.UserName;
                                            else
                                                cCheckBy1 = cCheckBy1 + "/" + rd.UserName;
                                        }
                                        if (rd.UDesc.Equals("Check"))
                                        {
                                            if (cCheckBy2.Equals(""))
                                                cCheckBy2 = rd.UserName;
                                            else
                                                cCheckBy2 = cCheckBy2 + "/" + rd.UserName;
                                        }
                                    }
                                    else
                                    {
                                        if (rd.UDesc.Equals("ผู้ตรวจสอบ"))
                                        {

                                            if (cCheckBy3.Equals(""))
                                                cCheckBy3 = rd.UserName;
                                            else
                                                cCheckBy3 = cCheckBy3 + "/" + rd.UserName;                                            
                                        }
                                        if (rd.UDesc.Equals("Check"))
                                        {
                                            if (cCheckBy4.Equals(""))
                                                cCheckBy4 = rd.UserName;
                                            else
                                                cCheckBy4 = cCheckBy4 + "/" + rd.UserName;
                                        }
                                    }
                                }


                                FormISO = qh.FormISO;
                                QHNo = qh.QCNo;
                                RefValue1 = qh.RefValue1;
                                RefValue2 = qh.RefValue2;
                                RefValue3 = qh.RefValue3;
                                string E5 = "";
                                int RowS = 6;
                                var GroupPartList = db.tb_QCGroupParts.Where(p => p.PartNo.Equals(PartNo) && p.FormISO.Equals(FromIS) && p.GroupPart.Trim().Equals("")).OrderBy(s => s.Seq).ToList();

                                if (PAGE1)
                                {
                                    InsertToExcel(ref worksheet, "AJ4", db.QC_GetUserName(qh.ApproveBy));
                                    InsertToExcel(ref worksheet, "AT5", cCheckBy1);
                                    InsertToExcel(ref worksheet, "AW5", cCheckBy3);
                                    //Step ตรวจสอบ//                                   
                                    //End Step//
                                    E5 = Convert.ToDateTime(qh.CreateDate).ToString("dd") + " วัน " + Convert.ToDateTime(qh.CreateDate).ToString("MM") + " เดือน  " + Convert.ToDateTime(qh.CreateDate).ToString("yyyy") + " ปี";
                                    InsertToExcel(ref worksheet, "K5", E5);
                                    InsertToExcel(ref worksheet, "J4", qh.OrderQty); //Prod. Qty
                                    InsertToExcel(ref worksheet, "K4", db.get_QCSumQtyNG_RE(qh.QCNo, 1)); // NG
                                    InsertToExcel(ref worksheet, "M4", db.get_QCSumQtyNG_RE(qh.QCNo, 2)); //Rework
                                    RowS = 6;
                                    foreach (var rd in GroupPartList)
                                    {
                                        RowS += 1;
                                        InsertToExcel(ref worksheet, "A" + (RowS).ToString(), rd.Seq);
                                        InsertToExcel(ref worksheet, "B" + (RowS).ToString(), rd.SetData);                                       
                                        if (rd.SetData.ToUpper().Equals("LOT"))
                                        {                                           
                                            InsertToExcel(ref worksheet, "B" + (RowS).ToString(), LotMark);
                                        }                                        
                                    }
                                    //LOT ต้องถูกต้อง, มีความชัดเจนสามารถอ่านได้ ( )
                                    LotMark = "LOT ต้องถูกต้อง, มีความชัดเจนสามารถอ่านได้ ( "+DValue.LotNo+" )";
                                    InsertToExcel(ref worksheet, "B14", LotMark);
                                    if (!TypeReport.Equals("HINO"))
                                    {
                                        InsertToExcel(ref worksheet, "A18", "");
                                        InsertToExcel(ref worksheet, "A19", "");
                                        InsertToExcel(ref worksheet, "B18", "");
                                        InsertToExcel(ref worksheet, "B19", "");
                                    }
                                }

                                if (PAGE2)
                                {
                                    InsertToExcel(ref worksheet2, "AJ4", db.QC_GetUserName(qh.ApproveBy));
                                    InsertToExcel(ref worksheet2, "AT5", cCheckBy1);
                                    InsertToExcel(ref worksheet2, "AW5", cCheckBy3);
                                    //Step ตรวจสอบ//                                   
                                    //End Step//
                                    E5 = Convert.ToDateTime(qh.CreateDate).ToString("dd") + " วัน " + Convert.ToDateTime(qh.CreateDate).ToString("MM") + " เดือน  " + Convert.ToDateTime(qh.CreateDate).ToString("yyyy") + " ปี";
                                    InsertToExcel(ref worksheet2, "K5", E5);
                                    InsertToExcel(ref worksheet2, "J4", qh.OrderQty); //Prod. Qty
                                    InsertToExcel(ref worksheet2, "K4", db.get_QCSumQtyNG_RE(qh.QCNo, 1)); // NG
                                    InsertToExcel(ref worksheet2, "M4", db.get_QCSumQtyNG_RE(qh.QCNo, 2)); //Rework
                                    RowS = 6;
                                    foreach (var rd in GroupPartList)
                                    {
                                        RowS += 1;
                                        InsertToExcel(ref worksheet2, "A" + (RowS).ToString(), rd.Seq);
                                        InsertToExcel(ref worksheet2, "B" + (RowS).ToString(), rd.SetData);
                                        if (rd.SetData.ToUpper().Equals("LOT"))
                                        {
                                            InsertToExcel(ref worksheet2, "B" + (RowS).ToString(), LotMark);
                                        }
                                    }

                                    LotMark = "LOT ต้องถูกต้อง, มีความชัดเจนสามารถอ่านได้ ( " + DValue.LotNo + " )";
                                    InsertToExcel(ref worksheet2, "B14", LotMark);
                                    if (!TypeReport.Equals("HINO"))
                                    {
                                        InsertToExcel(ref worksheet2, "A18", "");
                                        InsertToExcel(ref worksheet2, "A19", "");
                                        InsertToExcel(ref worksheet2, "B18", "");
                                        InsertToExcel(ref worksheet2, "B19", "");
                                    }
                                }
                                if (PAGE3)
                                {
                                    InsertToExcel(ref worksheet3, "AJ4", db.QC_GetUserName(qh.ApproveBy));
                                    InsertToExcel(ref worksheet3, "AT5", cCheckBy1);
                                    InsertToExcel(ref worksheet3, "AW5", cCheckBy3);
                                    //Step ตรวจสอบ//                                   
                                    //End Step//
                                    E5 = Convert.ToDateTime(qh.CreateDate).ToString("dd") + " วัน " + Convert.ToDateTime(qh.CreateDate).ToString("MM") + " เดือน  " + Convert.ToDateTime(qh.CreateDate).ToString("yyyy") + " ปี";
                                    InsertToExcel(ref worksheet3, "K5", E5);
                                    InsertToExcel(ref worksheet3, "J4", qh.OrderQty); //Prod. Qty
                                    InsertToExcel(ref worksheet3, "K4", db.get_QCSumQtyNG_RE(qh.QCNo, 1)); // NG
                                    InsertToExcel(ref worksheet3, "M4", db.get_QCSumQtyNG_RE(qh.QCNo, 2)); //Rework
                                    RowS = 6;
                                    foreach (var rd in GroupPartList)
                                    {
                                        RowS += 1;
                                        InsertToExcel(ref worksheet3, "A" + (RowS).ToString(), rd.Seq);
                                        InsertToExcel(ref worksheet3, "B" + (RowS).ToString(), rd.SetData);
                                        if (rd.SetData.ToUpper().Equals("LOT"))
                                        {
                                            InsertToExcel(ref worksheet3, "B" + (RowS).ToString(), LotMark);
                                        }
                                    }
                                    LotMark = "LOT ต้องถูกต้อง, มีความชัดเจนสามารถอ่านได้ ( " + DValue.LotNo + " )";
                                    InsertToExcel(ref worksheet3, "B14", LotMark);
                                    if (!TypeReport.Equals("HINO"))
                                    {
                                        InsertToExcel(ref worksheet3, "A18", "");
                                        InsertToExcel(ref worksheet3, "A19", "");
                                        InsertToExcel(ref worksheet3, "B18", "");
                                        InsertToExcel(ref worksheet3, "B19", "");
                                    }
                                }
                                if (PAGE4)
                                {
                                    InsertToExcel(ref worksheet4, "AJ4", db.QC_GetUserName(qh.ApproveBy));
                                    InsertToExcel(ref worksheet4, "AT5", cCheckBy1);
                                    InsertToExcel(ref worksheet4, "AW5", cCheckBy3);
                                    //Step ตรวจสอบ//                                   
                                    //End Step//
                                    E5 = Convert.ToDateTime(qh.CreateDate).ToString("dd") + " วัน " + Convert.ToDateTime(qh.CreateDate).ToString("MM") + " เดือน  " + Convert.ToDateTime(qh.CreateDate).ToString("yyyy") + " ปี";
                                    InsertToExcel(ref worksheet4, "K5", E5);
                                    InsertToExcel(ref worksheet4, "J4", qh.OrderQty); //Prod. Qty
                                    InsertToExcel(ref worksheet4, "K4", db.get_QCSumQtyNG_RE(qh.QCNo, 1)); // NG
                                    InsertToExcel(ref worksheet4, "M4", db.get_QCSumQtyNG_RE(qh.QCNo, 2)); //Rework
                                    RowS = 6;
                                    foreach (var rd in GroupPartList)
                                    {
                                        RowS += 1;
                                        InsertToExcel(ref worksheet4, "A" + (RowS).ToString(), rd.Seq);
                                        InsertToExcel(ref worksheet4, "B" + (RowS).ToString(), rd.SetData);
                                        if (rd.SetData.ToUpper().Equals("LOT"))
                                        {
                                            InsertToExcel(ref worksheet4, "B" + (RowS).ToString(), LotMark);
                                        }
                                    }
                                    LotMark = "LOT ต้องถูกต้อง, มีความชัดเจนสามารถอ่านได้ ( " + DValue.LotNo + " )";
                                    InsertToExcel(ref worksheet4, "B14", LotMark);
                                    if (!TypeReport.Equals("HINO"))
                                    {
                                        InsertToExcel(ref worksheet4, "A18", "");
                                        InsertToExcel(ref worksheet4, "A19", "");
                                        InsertToExcel(ref worksheet4, "B18", "");
                                        InsertToExcel(ref worksheet4, "B19", "");
                                    }
                                }
                            }

                        }
                        catch (Exception ex) { MessageBox.Show("first " + ex.Message); }
                    }

                    ////////////////////////////////////////
                    int SetCol = 10;
                    int countA = 0;
                    int CountB = 0;
                    int CountC = 0;
                    int CountD = 0;
                    int TAG2 = 0;
                    int CA = 0;
                    int TG2 = 0;
                    int NGA = 0;
                    int NGB = 0;
                    int NGC = 0;
                    string TAGOf1 = "";
                    string TAGOf2 = "";
                    string TAGOf3 = "";

                    int CountTAG = 0;
                    var listPoint = db.sp_46_QCSelectWO_09_QCTAGSelect(QHNo).ToList();
                    CountTAG = listPoint.Count;
                    if (listPoint.Count > 0)
                    {
                        foreach (var rs in listPoint)
                        {
                            countA += 1;
                            if (countA > 40)
                            {
                                CountB += 1;
                            }
                            if (countA > 80)
                            {
                                CountC += 1;
                            }
                            if (countA > 120)
                            {
                                CountD += 1;
                            }
                            TG = 0;

                            string[] PPTAG = rs.BarcodeTag.Split(',');
                            TG = Convert.ToInt32(PPTAG[2]);
                            TAG2 += TG;
                            TG2 = 0;
                            TG2 = Convert.ToInt32(db.get_QCSumQtyTAGNG(QHNo, rs.BarcodeTag, 3));

                            if (listPoint.Count == countA)
                            {
                                NGA = TG;
                                TAGOf1 = PPTAG[5];
                            }
                            if ((listPoint.Count - 1) == countA)
                            {
                                NGB = TG;
                                TAGOf2 = PPTAG[5];
                            }
                            if ((listPoint.Count - 2) == countA)
                            {
                                NGC = TG;
                                TAGOf3 = PPTAG[5];
                            }
                            row1 = 6;
                            Seq = 0;
                           
                            var listPart = db.tb_QCGroupParts.Where(q => q.FormISO.Equals(FromIS) && q.PartNo.Equals(DValue.CODE) && q.GroupPart.Trim().Equals("")).OrderBy(o => o.Seq).ToList();
                            CA = listPart.Count();
                            foreach (var rd in listPart)
                            {
                                row1 += 1;
                                Seq += 1;
                                if (!rd.SetData.Equals("") && row1 <= 19)
                                {
                                    try
                                    {
                                        var gValue = db.sp_46_QCGetValue5601(rs.BarcodeTag, QHNo, rd.Seq).FirstOrDefault();
                                        PV = "P";
                                        if (gValue.CountA > 0)
                                        {
                                            PV = "O";

                                            if (gValue.CountA == 99)
                                            {
                                                PV = "";
                                            }
                                        }
                                        //if (rd.Seq.Equals(9) && C9.Equals(""))
                                        //{
                                        //    PV = "";
                                        //}
                                        if (countA <= 40)
                                        {
                                            InsertToExcel(ref worksheet, Getcolumn(countA + SetCol) + row1.ToString(), PV);
                                        }
                                        if (countA > 40 && countA <= 80)
                                        {
                                            InsertToExcel(ref worksheet2, Getcolumn(CountB + SetCol) + row1.ToString(), PV);
                                        }
                                        if (countA > 80 && countA <= 120)
                                        {
                                            InsertToExcel(ref worksheet3, Getcolumn(CountC + SetCol) + row1.ToString(), PV);
                                        }
                                        if (countA > 120 && countA <= 160)
                                        {
                                            InsertToExcel(ref worksheet4, Getcolumn(CountD + SetCol) + row1.ToString(), PV);
                                        }
                                    }
                                    catch { }
                                    //catch (Exception ex) { MessageBox.Show("Mid " + ex.Message); }

                                }
                                //SumNG//       

                                if (countA <= 40)
                                {
                                    InsertToExcel(ref worksheet, Getcolumn(countA + SetCol) + "21", (TG - TG2));
                                    InsertToExcel(ref worksheet, Getcolumn(countA + SetCol) + "22", TG2);
                                }
                                else if (countA > 40 && countA <= 80)
                                {
                                    if (PAGE2)
                                    {
                                        InsertToExcel(ref worksheet2, Getcolumn(CountB + SetCol) + "21", (TG - TG2));
                                        InsertToExcel(ref worksheet2, Getcolumn(CountB + SetCol) + "22", TG2);
                                    }
                                }
                                else if (countA > 80 && countA <= 120)
                                {
                                    if (PAGE3)
                                    {
                                        InsertToExcel(ref worksheet3, Getcolumn(CountC + SetCol) + "21", (TG - TG2));
                                        InsertToExcel(ref worksheet3, Getcolumn(CountC + SetCol) + "22", TG2);

                                    }
                                }
                                else if (countA > 120 && countA <= 160)
                                {
                                    if (PAGE4)
                                    {
                                        InsertToExcel(ref worksheet4, Getcolumn(CountD + SetCol) + "21", (TG - TG2));
                                        InsertToExcel(ref worksheet4, Getcolumn(CountD + SetCol) + "22", TG2);
                                    }
                                }

                            }//foreach 
                             //}//cunt A //Page 1 End
                        }//for

                        ////NGQty and Remark//
                        int RM = 0;
                        int TNG1 = 0;
                        int TNG2 = 0;
                        int TNG3 = 0;
                        tb_QCProblem qcp = db.tb_QCProblems.Where(p => p.QCNo.Equals(QHNo) && !p.NGQty.Equals(0)).FirstOrDefault();
                        if (qcp != null)
                        {
                            var tgf = db.tb_QCTAGs.Where(s => s.QCNo.Equals(QHNo)).ToList();
                            foreach (var tf in tgf)
                            {
                                if (tf.ofTAG.Equals(TAGOf1))
                                {
                                    TNG1 = Convert.ToInt32(tf.NGQty);
                                }
                                else if (tf.ofTAG.Equals(TAGOf2))
                                {
                                    TNG2 = Convert.ToInt32(tf.NGQty);
                                }
                                else if (tf.ofTAG.Equals(TAGOf3))
                                {
                                    TNG3 = Convert.ToInt32(tf.NGQty);
                                }
                            }

                            if (countA <= 40)
                            {   //NG Qty//                
                                InsertToExcel(ref worksheet, "B20", qcp.ProblemName);
                                ///////////////////////////////
                                if (qcp.NGQty > 0)
                                {
                                    if (NGA < (qcp.NGQty + TNG1))
                                    {
                                        InsertToExcel(ref worksheet, Getcolumn(countA + SetCol) + "20", "O");
                                        InsertToExcel(ref worksheet, Getcolumn(countA + SetCol) + "21", 0);
                                        InsertToExcel(ref worksheet, Getcolumn(countA + SetCol) + "22", NGA);
                                        RM = (Convert.ToInt32(qcp.NGQty) + TNG1) - NGA;
                                        if (RM > 0)
                                        {
                                            if (NGB < (RM + TNG2))
                                            {
                                                InsertToExcel(ref worksheet, Getcolumn((countA - 1) + SetCol) + "20", "O");
                                                InsertToExcel(ref worksheet, Getcolumn((countA - 1) + SetCol) + "21", 0);
                                                InsertToExcel(ref worksheet, Getcolumn((countA - 1) + SetCol) + "22", NGB);
                                                RM = (RM + TNG2) - NGB;
                                                if (RM > 0)
                                                {
                                                    if (NGC < (RM + TNG3))
                                                    {
                                                        InsertToExcel(ref worksheet, Getcolumn((countA - 2) + SetCol) + "20", "O");
                                                        InsertToExcel(ref worksheet, Getcolumn((countA - 2) + SetCol) + "21", 0);
                                                        InsertToExcel(ref worksheet, Getcolumn((countA - 2) + SetCol) + "22", NGC);
                                                    }
                                                    else
                                                    {
                                                        InsertToExcel(ref worksheet, Getcolumn((countA - 2) + SetCol) + "20", "O");
                                                        InsertToExcel(ref worksheet, Getcolumn((countA - 2) + SetCol) + "21", NGC - (RM + TNG3));
                                                        InsertToExcel(ref worksheet, Getcolumn((countA - 2) + SetCol) + "22", RM + TNG3);
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                InsertToExcel(ref worksheet, Getcolumn((countA - 1) + SetCol) + "20", "O");
                                                InsertToExcel(ref worksheet, Getcolumn((countA - 1) + SetCol) + "21", NGB - (RM + TNG2));
                                                InsertToExcel(ref worksheet, Getcolumn((countA - 1) + SetCol) + "22", (RM + TNG2));
                                            }
                                        }


                                    }
                                    else
                                    {
                                        InsertToExcel(ref worksheet, Getcolumn((countA - 0) + SetCol) + "20", "O");
                                        InsertToExcel(ref worksheet, Getcolumn((countA - 0) + SetCol) + "21", NGA - (qcp.NGQty + TNG1));
                                        InsertToExcel(ref worksheet, Getcolumn((countA - 0) + SetCol) + "22", (qcp.NGQty + TNG1));

                                    }

                                }

                            }
                            else if (countA > 40 && countA <= 80)
                            {
                                if (PAGE2)
                                {
                                    //NG Qty//                                   
                                    InsertToExcel(ref worksheet2, "B20", qcp.ProblemName);
                                    if (qcp.NGQty > 0)
                                    {
                                        if (NGA < (qcp.NGQty + TNG1))
                                        {
                                            InsertToExcel(ref worksheet2, Getcolumn(countA + SetCol) + "20", "O");
                                            InsertToExcel(ref worksheet2, Getcolumn(countA + SetCol) + "21", 0);
                                            InsertToExcel(ref worksheet2, Getcolumn(countA + SetCol) + "22", NGA);
                                            RM = (Convert.ToInt32(qcp.NGQty) + TNG1) - NGA;
                                            if (RM > 0)
                                            {
                                                if (NGB < (RM + TNG2))
                                                {

                                                    InsertToExcel(ref worksheet2, Getcolumn((countA - 1) + SetCol) + "20", "O");
                                                    InsertToExcel(ref worksheet2, Getcolumn((countA - 1) + SetCol) + "21", 0);
                                                    InsertToExcel(ref worksheet2, Getcolumn((countA - 1) + SetCol) + "22", NGB);

                                                    RM = (RM + TNG2) - NGB;
                                                    if (RM > 0)
                                                    {
                                                        if (NGC < (RM + TNG3))
                                                        {

                                                            InsertToExcel(ref worksheet2, Getcolumn((countA - 2) + SetCol) + "20", "O");
                                                            InsertToExcel(ref worksheet2, Getcolumn((countA - 2) + SetCol) + "21", 0);
                                                            InsertToExcel(ref worksheet2, Getcolumn((countA - 2) + SetCol) + "22", NGC);
                                                        }
                                                        else
                                                        {
                                                            InsertToExcel(ref worksheet2, Getcolumn((countA - 2) + SetCol) + "20", "O");
                                                            InsertToExcel(ref worksheet2, Getcolumn((countA - 2) + SetCol) + "21", NGC - (RM + TNG3));
                                                            InsertToExcel(ref worksheet2, Getcolumn((countA - 2) + SetCol) + "22", RM + TNG3);
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    InsertToExcel(ref worksheet2, Getcolumn((countA - 1) + SetCol) + "20", "O");
                                                    InsertToExcel(ref worksheet2, Getcolumn((countA - 1) + SetCol) + "21", NGB - (RM + TNG2));
                                                    InsertToExcel(ref worksheet2, Getcolumn((countA - 1) + SetCol) + "22", (RM + TNG2));
                                                }
                                            }


                                        }
                                        else
                                        {
                                            InsertToExcel(ref worksheet2, Getcolumn((countA - 0) + SetCol) + "20", "O");
                                            InsertToExcel(ref worksheet2, Getcolumn((countA - 0) + SetCol) + "21", NGA - (qcp.NGQty + TNG1));
                                            InsertToExcel(ref worksheet2, Getcolumn((countA - 0) + SetCol) + "22", (qcp.NGQty + TNG1));
                                        }

                                    }
                                }
                            }
                            else if (countA > 80 && countA <= 120)
                            {
                                if (PAGE3)
                                {
                                    //NG Qty//                                   
                                    InsertToExcel(ref worksheet3, "B20", qcp.ProblemName);
                                    if (qcp.NGQty > 0)
                                    {
                                        if (NGA < (qcp.NGQty + TNG1))
                                        {
                                            InsertToExcel(ref worksheet3, Getcolumn(countA + SetCol) + "20", "O");
                                            InsertToExcel(ref worksheet3, Getcolumn(countA + SetCol) + "21", 0);
                                            InsertToExcel(ref worksheet3, Getcolumn(countA + SetCol) + "22", NGA);
                                            RM = (Convert.ToInt32(qcp.NGQty) + TNG1) - NGA;
                                            if (RM > 0)
                                            {
                                                if (NGB < (RM + TNG2))
                                                {

                                                    InsertToExcel(ref worksheet3, Getcolumn((countA - 1) + SetCol) + "20", "O");
                                                    InsertToExcel(ref worksheet3, Getcolumn((countA - 1) + SetCol) + "21", 0);
                                                    InsertToExcel(ref worksheet3, Getcolumn((countA - 1) + SetCol) + "22", NGB);
                                                    RM = (RM + TNG2) - NGB;
                                                    if (RM > 0)
                                                    {
                                                        if (NGC < (RM + TNG3))
                                                        {
                                                            InsertToExcel(ref worksheet3, Getcolumn((countA - 2) + SetCol) + "20", "O");
                                                            InsertToExcel(ref worksheet3, Getcolumn((countA - 2) + SetCol) + "21", 0);
                                                            InsertToExcel(ref worksheet3, Getcolumn((countA - 2) + SetCol) + "22", NGC);
                                                        }
                                                        else
                                                        {
                                                            InsertToExcel(ref worksheet3, Getcolumn((countA - 2) + SetCol) + "20", "O");
                                                            InsertToExcel(ref worksheet3, Getcolumn((countA - 2) + SetCol) + "21", NGC - (RM + TNG3));
                                                            InsertToExcel(ref worksheet3, Getcolumn((countA - 2) + SetCol) + "22", RM + TNG3);
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    InsertToExcel(ref worksheet3, Getcolumn((countA - 1) + SetCol) + "20", "O");
                                                    InsertToExcel(ref worksheet3, Getcolumn((countA - 1) + SetCol) + "21", NGB - (RM + TNG2));
                                                    InsertToExcel(ref worksheet3, Getcolumn((countA - 1) + SetCol) + "22", (RM + TNG2));
                                                }
                                            }


                                        }
                                        else
                                        {
                                            InsertToExcel(ref worksheet3, Getcolumn((countA - 0) + SetCol) + "20", "O");
                                            InsertToExcel(ref worksheet3, Getcolumn((countA - 0) + SetCol) + "21", NGA - (qcp.NGQty + TNG1));
                                            InsertToExcel(ref worksheet3, Getcolumn((countA - 0) + SetCol) + "22", (qcp.NGQty + TNG1));
                                        }

                                    }
                                }
                            }
                            else if (countA > 120 && countA <= 160)
                            {
                                if (PAGE4)
                                {
                                    //NG Qty//                                    
                                    InsertToExcel(ref worksheet4, "BC20", qcp.ProblemName);

                                    if (qcp.NGQty > 0)
                                    {
                                        if (NGA < (qcp.NGQty + TNG1))
                                        {
                                            InsertToExcel(ref worksheet4, Getcolumn(countA + SetCol) + "20", "O");
                                            InsertToExcel(ref worksheet4, Getcolumn(countA + SetCol) + "21", 0);
                                            InsertToExcel(ref worksheet4, Getcolumn(countA + SetCol) + "22", NGA);
                                            RM = (Convert.ToInt32(qcp.NGQty) + TNG1) - NGA;
                                            if (RM > 0)
                                            {
                                                if (NGB < (RM + TNG2))
                                                {
                                                    InsertToExcel(ref worksheet4, Getcolumn((countA - 1) + SetCol) + "20", "O");
                                                    InsertToExcel(ref worksheet4, Getcolumn((countA - 1) + SetCol) + "21", 0);
                                                    InsertToExcel(ref worksheet4, Getcolumn((countA - 1) + SetCol) + "22", NGB);
                                                    RM = (RM + TNG2) - NGB;
                                                    if (RM > 0)
                                                    {
                                                        if (NGC < (RM + TNG3))
                                                        {
                                                            InsertToExcel(ref worksheet4, Getcolumn((countA - 2) + SetCol) + "20", "O");
                                                            InsertToExcel(ref worksheet4, Getcolumn((countA - 2) + SetCol) + "21", 0);
                                                            InsertToExcel(ref worksheet4, Getcolumn((countA - 2) + SetCol) + "22", NGC);
                                                        }
                                                        else
                                                        {
                                                            InsertToExcel(ref worksheet4, Getcolumn((countA - 2) + SetCol) + "20", "O");
                                                            InsertToExcel(ref worksheet4, Getcolumn((countA - 2) + SetCol) + "21", NGC - (RM + TNG3));
                                                            InsertToExcel(ref worksheet4, Getcolumn((countA - 2) + SetCol) + "22", RM + TNG3);
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    InsertToExcel(ref worksheet4, Getcolumn((countA - 1) + SetCol) + "20", "O");
                                                    InsertToExcel(ref worksheet4, Getcolumn((countA - 1) + SetCol) + "21", NGB - (RM + TNG2));
                                                    InsertToExcel(ref worksheet4, Getcolumn((countA - 1) + SetCol) + "22", (RM + TNG2));
                                                }
                                            }


                                        }
                                        else
                                        {
                                            InsertToExcel(ref worksheet4, Getcolumn((countA - 0) + SetCol) + "20", "O");
                                            InsertToExcel(ref worksheet4, Getcolumn((countA - 0) + SetCol) + "21", NGA - (qcp.NGQty + TNG1));
                                            InsertToExcel(ref worksheet4, Getcolumn((countA - 0) + SetCol) + "22", (qcp.NGQty + TNG1));
                                        }

                                    }
                                }
                            }
                        }

                        ////// PC Check ///
                    }



                }

                excelBook.SaveAs(tempfile);
                excelBook.Close(false);
                excelApp.Quit();
                releaseObject(worksheet);
                releaseObject(worksheet2);
                releaseObject(worksheet3);
                releaseObject(excelBook);
                releaseObject(excelApp);

                Marshal.FinalReleaseComObject(worksheet);
                Marshal.FinalReleaseComObject(worksheet2);
                Marshal.FinalReleaseComObject(worksheet3);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet2);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet3);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                GC.GetTotalMemory(false);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.GetTotalMemory(true);
                System.Diagnostics.Process.Start(tempfile);

            }
            catch (Exception ex) { MessageBox.Show("last " + ex.Message); }
        }

        public static void PrintFMPD109(string WO, string PartNo, string QCNo1, string FromIS)
        {
            try
            {
                try
                {
                    string DATA = AppDomain.CurrentDomain.BaseDirectory;
                    string tempPath = System.IO.Path.GetTempPath();
                    string HINO = "";
                    string FileName = "FM-PD-109.xlsx";
                    FromIS = "FM-PD-109";


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
                    //int row1 = 22;
                    //int row2 = 22;
                    //int Seq = 0;
                    //int seq2 = 22;
                    //int CountRow = 0;
                    string cIssueBy1 = "";
                    string cIssueBy2 = "";
                    string cIssueBy3 = "";
                    string cIssueBy4 = "";

                    string cCheckBy1 = "";
                    string cCheckBy2 = "";
                    string cCheckBy3 = "";

                    string cCheckByF1 = "";
                    string cCheckByF2 = "";
                    string cCheckByF3 = "";

                    string PV = "P";
                    string QHNo = QCNo1;
                    string FormISO = FromIS;
                    string DN = "";
                    string SymBo = "～";
                    using (DataClasses1DataContext db = new DataClasses1DataContext())
                    {
                        //string Value1 = "";
                        //string Value2 = "";
                        //string LotNo = "";

                        ///////////////SETValue/////////////////
                        var DValue = db.sp_46_QCSelectWO_01(WO).FirstOrDefault();
                        if (DValue != null)
                        {
                            DN = DValue.DayNight;

                            InsertToExcel(ref worksheet, "I5", GetDWG(PartNo, 1));
                            InsertToExcel(ref worksheet, "AF1", GetDWG(PartNo, 2));
                            
                            InsertToExcel(ref worksheet, "P3", DValue.CODE.ToString());
                            InsertToExcel(ref worksheet, "P5", DValue.NAME.ToString());
                            //InsertToExcel(ref worksheet, "D3", "TW02-SC_PB");
                            InsertToExcel(ref worksheet, "D5", DValue.PORDER.ToString());
                            InsertToExcel(ref worksheet, "D7", DValue.DeliveryDate);
                            InsertToExcel(ref worksheet, "D9", DValue.LotNo.ToString());
                            InsertToExcel(ref worksheet, "D11", DValue.OrderQty.ToString());
                            var gTime = db.sp_46_QCGetValue2601_Time(WO).ToList();
                            if (gTime.Count > 0)
                            {
                                var g = gTime.FirstOrDefault();
                                DateTime Chtime = Convert.ToDateTime(g.BomTime);
                                DateTime Chtime2 = Convert.ToDateTime(g.PrintTime);

                                if (g.BomTime2 != "")
                                    Chtime = Convert.ToDateTime(g.BomTime2);
                                if (g.BomTime == g.PrintTime)
                                {
                                    Chtime2 = Convert.ToDateTime(g.PrintTime).AddMinutes(30);
                                }
                                InsertToExcel(ref worksheet, "AB9", Math.Abs(Convert.ToDecimal((Chtime - Chtime2).TotalMinutes)).ToString("####") + " นาที");
                                if (!g.StartTime.Equals(""))
                                {
                                    InsertToExcel(ref worksheet, "N7", Convert.ToDateTime(Chtime2).ToString("HH:mm"));
                                    InsertToExcel(ref worksheet, "AA7", Convert.ToDateTime(g.EndTime).ToString("HH:mm"));
                                    InsertToExcel(ref worksheet, "O9", "'" + Convert.ToDateTime(Chtime).ToString("HH:mm") + " " + SymBo + " " + Convert.ToDateTime(Chtime2).ToString("HH:mm"));
                                }
                            }
                            try
                            {
                                tb_QCHD qh = db.tb_QCHDs.Where(w => w.QCNo.Equals(QCNo1)).FirstOrDefault();
                                if (qh != null)
                                {
                                    //////////Find UserName////////////
                                    var uc = db.tb_QCCheckUsers.Where(u => u.QCNo.Equals(QCNo1)).ToList();
                                    int r1 = 0;
                                    int r2 = 0;
                                    int r3 = 0;
                                    int rr1 = 0;
                                    int rr2 = 0;
                                    int rr3 = 0;
                                    foreach (var rd in uc)
                                    {
                                        DN = rd.DayN;// dbShowData.CheckDayN(Convert.ToDateTime(rd.ScanDate));
                                        if (DN.Equals("D"))
                                        {
                                            if (rd.UDesc.Equals("ผู้จัดทำเอกสาร"))
                                                cIssueBy1 = rd.UserName;
                                            if (rd.UDesc.Equals("ผู้ตรวจสอบก่อนผลิต"))
                                                cIssueBy2 = rd.UserName;
                                        }
                                        else //N
                                        {
                                            if (rd.UDesc.Equals("ผู้จัดทำเอกสาร"))
                                                cIssueBy3 = rd.UserName;
                                            if (rd.UDesc.Equals("ผู้ตรวจสอบก่อนผลิต"))
                                                cIssueBy4 = rd.UserName;
                                        }
                                    }

                                    InsertToExcel(ref worksheet, "AE10", db.QC_GetUserName(qh.ApproveBy));
                                    InsertToExcel(ref worksheet, "AE5", "1. " + cIssueBy1);
                                    InsertToExcel(ref worksheet, "AE7", "2. " + cIssueBy2);
                                    InsertToExcel(ref worksheet, "AF5", "1. " + cIssueBy3);
                                    InsertToExcel(ref worksheet, "AF7", "2. " + cIssueBy4);

                                    //Main Line
                                    InsertToExcel(ref worksheet, "E23", db.get_QC_UserNameScan(qh.QCNo, "Sub line", "D"));
                                    InsertToExcel(ref worksheet, "F23", db.get_QC_UserNameScan(qh.QCNo, "Sub line", "N"));

                                    InsertToExcel(ref worksheet, "E32", db.get_QC_UserNameScan(qh.QCNo, "Main line", "D"));
                                    InsertToExcel(ref worksheet, "F32", db.get_QC_UserNameScan(qh.QCNo, "Main line", "N"));

                                    InsertToExcel(ref worksheet, "E40", db.get_QC_UserNameScan(qh.QCNo, "ตรวจสอบท้ายไลน์", "D"));
                                    InsertToExcel(ref worksheet, "F40", db.get_QC_UserNameScan(qh.QCNo, "ตรวจสอบท้ายไลน์", "N"));


                                    //ตอก Lot//
                                    InsertToExcel(ref worksheet, "AE36", "LOT ( " + qh.LotNo + " )");
                                    InsertToExcel(ref worksheet, "AE37", DValue.DayNight);
                                    // InsertToExcel(ref worksheet, "AE33", db.get_QC_DATAPoint_AG(qh.WONo, 35));
                                    //    InsertToExcel(ref worksheet, "AE34", db.get_QC_DATAPoint_AG(qh.WONo, 36));


                                    //Find CountPD
                                    /////////////////////////////
                                    int cpd = 0;
                                    int sss = 0;
                                    /////////////////////////////
                                    var co = db.tb_QCCountPDs.Where(c => c.WONo.Equals(WO)).ToList();
                                    foreach (var rd in co)
                                    {
                                        cpd += 1;

                                        if (cpd <= 5)
                                        {
                                            if (rd.DayN.Equals("D"))
                                            {
                                                InsertToExcel(ref worksheet, "F" + (46 + cpd).ToString(), rd.A1);
                                            }
                                            else
                                            {
                                                InsertToExcel(ref worksheet, "H" + (46 + cpd).ToString(), rd.A1);
                                            }
                                        }
                                        else
                                        {
                                            if (rd.DayN.Equals("D"))
                                            {
                                                InsertToExcel(ref worksheet, "R" + (46 + (cpd - 5)).ToString(), rd.A1);
                                            }
                                            else
                                            {
                                                InsertToExcel(ref worksheet, "T" + (46 + (cpd - 5)).ToString(), rd.A1);
                                            }
                                        }

                                    }
                                    //Find Problem//

                                    tb_QCProblem pb = db.tb_QCProblems.Where(p => p.QCNo.Equals(QHNo)).FirstOrDefault();
                                    if (pb != null)
                                    {
                                        if (pb.TypeProblem.Equals("Man"))
                                        {
                                            InsertToExcel(ref worksheet, "F13", "P");
                                        }
                                        else if (pb.TypeProblem.Equals("Machine"))
                                        {
                                            InsertToExcel(ref worksheet, "F13", "P");
                                        }
                                        else if (pb.TypeProblem.Equals("Method"))
                                        {
                                            InsertToExcel(ref worksheet, "M13", "P");
                                        }
                                        else if (pb.TypeProblem.Equals("Material"))
                                        {
                                            InsertToExcel(ref worksheet, "P13", "P");
                                        }
                                        else if (pb.TypeProblem.Equals("Other"))
                                        {
                                            InsertToExcel(ref worksheet, "S13", "P");
                                            InsertToExcel(ref worksheet, "X13", pb.TypeRemark);
                                        }
                                        InsertToExcel(ref worksheet, "E14", pb.ProblemSeeBy);
                                        InsertToExcel(ref worksheet, "N14", pb.ProblemName);
                                        InsertToExcel(ref worksheet, "AC14", pb.ProblemWare);
                                        InsertToExcel(ref worksheet, "E15", pb.ProblemTime);
                                        InsertToExcel(ref worksheet, "N15", pb.ProblemWhy);
                                        InsertToExcel(ref worksheet, "E17", pb.ProblemFix);
                                        InsertToExcel(ref worksheet, "V18", pb.FixBy);
                                        InsertToExcel(ref worksheet, "AE18", pb.CheckBy);

                                    }
                                    ////Scan Time///                        

                                    ////Set Topic//

                                    InsertToExcel(ref worksheet, "L42", db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 50));
                                    InsertToExcel(ref worksheet, "L43", db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 51));
                                    InsertToExcel(ref worksheet, "L44", db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 52));
                                    //Step 1
                                    int cRow = 22;                                   
                                    string Ppart = "";
                                    string Pparg2 = "";
                                    var QCP = db.sp_46_QCSelectWO_13_GroupPart(qh.WONo).ToList();
                                    cRow = 22;
                                    foreach (var rx in QCP)
                                    {
                                        Ppart = "";
                                        Pparg2 = "";
                                        cRow += 1;
                                        if (cRow < 42)
                                        {                                                                                      
                                            Ppart = rx.PartNo;
                                            Pparg2 = db.getItemNo(rx.PartNo);
                                            if (rx.PartNo.ToUpper().Contains("GREASE") || rx.PartNo.ToUpper().Contains("SILICON") || rx.PartNo.ToUpper().Contains("LOCTITE"))
                                            {
                                                Pparg2 = "";
                                                                                                                                   
                                            }                                           
                                            if (Pparg2 != "")
                                            {
                                                InsertToExcel(ref worksheet, "G" + cRow.ToString(), Pparg2);
                                                InsertToExcel(ref worksheet, "L" + cRow.ToString(), Ppart);
                                                Ppart = rx.PartNo;
                                                var rds = db.sp_46_QCGetValue2601(qh.WONo, Ppart).FirstOrDefault();
                                                if (rds != null)
                                                {
                                                    InsertToExcel(ref worksheet, "Q" + cRow.ToString(), rds.DayN);
                                                    InsertToExcel(ref worksheet, "R" + cRow.ToString(), rds.NightN);
                                                    InsertToExcel(ref worksheet, "S" + cRow.ToString(), rds.Lot);
                                                }
                                            }
                                        }
                                    }

                                    InsertToExcel(ref worksheet, "G42", db.get_QC_SetDataMaster2(qh.FormISO, qh.PartNo, 50));
                                    InsertToExcel(ref worksheet, "G43", db.get_QC_SetDataMaster2(qh.FormISO, qh.PartNo, 51));
                                    InsertToExcel(ref worksheet, "G44", db.get_QC_SetDataMaster2(qh.FormISO, qh.PartNo, 52));

                                    InsertToExcel(ref worksheet, "L42", db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 50));
                                    InsertToExcel(ref worksheet, "L43", db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 51));
                                    InsertToExcel(ref worksheet, "L44", db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, 52));
                                    string ppA = "";
                                    int qrs = 41;
                                    for (int i = 0; i < 3; i++)
                                    {
                                        qrs += 1;
                                        ppA = db.get_QC_SetDataMaster2(qh.FormISO, qh.PartNo, 50+i);
                                        var rds = db.sp_46_QCGetValue2601(qh.WONo, ppA).FirstOrDefault();
                                        if (rds != null)
                                        {
                                            InsertToExcel(ref worksheet, "Q" + qrs.ToString(), rds.DayN);
                                            InsertToExcel(ref worksheet, "R" + qrs.ToString(), rds.NightN);
                                            InsertToExcel(ref worksheet, "S" + qrs.ToString(), rds.Lot);
                                        }
                                    }

                                    /////Step 2
                                    ////  int crow2 = 22;
                                    cRow = 22;
                                    int CK = 0;
                                    int N23 = 0;
                                    int D23 = 0;                                   

                                    for (int II = 1; II <= 17; II++)
                                    {

                                        cRow += 1; 
                                        Ppart = "";
                                        Pparg2 = "";
                                        Ppart = db.get_QC_SetDataMaster(qh.FormISO, qh.PartNo, II); //Set Data
                                        Pparg2 = db.get_QC_SetDataMaster2(qh.FormISO, qh.PartNo, II);//Topic                                      

                                        if (Ppart.Trim() != "")
                                        {
                                            InsertToExcel(ref worksheet, "AE" + cRow.ToString(), Ppart);
                                            // InsertToExcel(ref worksheet, "AA" + cRow.ToString(), Pparg2);
                                        }
                                        if (Ppart.ToUpper().Contains("HINO"))
                                        {
                                            HINO = "HINO";
                                        }                                     

                                        ////Line 1 //
                                        CK = Convert.ToInt32(db.get_QC_DATAPoint_AG_2D(qh.WONo, qh.PartNo, qh.FormISO, II));
                                        if (CK == 1)
                                        {
                                            InsertToExcel(ref worksheet, "AG" + cRow.ToString(), "P");
                                            D23 += 1;
                                        }
                                        else if (CK == 2)
                                        {
                                            InsertToExcel(ref worksheet, "AH" + cRow.ToString(), "P");
                                            N23 += 1;
                                        }
                                        else if (CK == 3)
                                        {
                                            D23 += 1;
                                            N23 += 1;                                            
                                            InsertToExcel(ref worksheet, "AG" + cRow.ToString(), "P");                                           
                                            InsertToExcel(ref worksheet, "AH" + cRow.ToString(), "P");                                            
                                        }

                                        if (II == 14)
                                        {
                                            if (D23 > 0)
                                            {
                                                InsertToExcel(ref worksheet, "AG" + cRow.ToString(), "P");
                                            }
                                            if (N23 > 23)
                                            {
                                                InsertToExcel(ref worksheet, "AH" + cRow.ToString(), "P");
                                            }
                                        }
                                    }
                                    //Update ISUZU//
                                    if (HINO.Equals(""))
                                    {
                                        InsertToExcel(ref worksheet, "Z39", ""); 
                                        InsertToExcel(ref worksheet, "AG39", "");
                                        InsertToExcel(ref worksheet, "AH39", "");
                                        InsertToExcel(ref worksheet, "AA39", "");
                                        InsertToExcel(ref worksheet, "AE39", "");
                                        InsertToExcel(ref worksheet, "N51", "");

                                    }
                                }
                                
                            }
                            catch { }
                        }
                        ////////////////////////////////////////
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
            }
            catch { }
        }
        public static void PrintFMPD003_S(string WO, string PartNo, string QCNo1, string FromIS)
        {
            try
            {
                string DATA = AppDomain.CurrentDomain.BaseDirectory;
                string tempPath = System.IO.Path.GetTempPath();
                string FileName = "FM-PD-156_S_HINO.xlsx";
                string tempfile = tempPath + FileName;
                string TypeReport = GetReportName("HINO", PartNo, FromIS).ToUpper();
                string HeaderText = " Shipping Check  result in process  (Piggy Back Piston Type)";             
                
                if (!TypeReport.Equals("HINO"))
                {
                    FileName = "FM-PD-156_S_ISUZU.xlsx";
                    tempfile = tempPath + FileName;
                }
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
                //if (!TypeReport.Equals("HINO"))
                //{                    
                //    worksheet = (Excel.Worksheet)sheets.get_Item(2);
                //}
                worksheet.Activate();

                // progressBar1.Maximum = 51;
                // progressBar1.Minimum = 1;
                int row1 = 14;

                string PV = "P";
                string QHNo = QCNo1;
                string FormISO = FromIS;
                string DN = "";
                string cCheckBy1 = "";
                string cCheckBy2 = "";
                string cCheckBy3 = "";
                string LotRemark = "";
                string E5 = "";
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    string PartName = "";
                    string LotNo = "";
                    ///////////////SETValue/////////////////
                    var DValue = db.sp_46_QCSelectWO_01(WO).FirstOrDefault();
                    if (DValue != null)
                    {
                        DN = DValue.DayNight;
                        PartName = DValue.NAME;
                        InsertToExcel(ref worksheet, "A4", HeaderText);
                        InsertToExcel(ref worksheet, "C6", DValue.NAME);
                        InsertToExcel(ref worksheet, "C7", DValue.CODE);
                        InsertToExcel(ref worksheet, "C8", db.get_QC_ItemCustomerNo(DValue.CODE, ""));
                        E5 = " วัน " + Convert.ToDateTime(DValue.CreateDate).ToString("dd") + " เดือน " + Convert.ToDateTime(DValue.CreateDate).ToString("MM") + " ปี  " + Convert.ToDateTime(DValue.CreateDate).ToString("yyyy") + "";
                        InsertToExcel(ref worksheet, "C9", E5);
                        InsertToExcel(ref worksheet, "C10", DValue.OrderQty);
                        InsertToExcel(ref worksheet, "C11", DValue.LotNo);
                        InsertToExcel(ref worksheet, "D13", DValue.PORDER);
                        LotRemark = "มีการตอก Lot no.ถูกต้อง,ชัดเจน,ตำแหน่งการตอกถูกต้อง ( " + DValue.LotNo + " ) ";
                        InsertToExcel(ref worksheet, "D22", LotRemark);
                        if (TypeReport.Equals("HINO"))
                        {
                            InsertToExcel(ref worksheet, "D26", db.get_QC_SetDataMaster(FromIS,DValue.CODE,12));
                            InsertToExcel(ref worksheet, "D27", db.get_QC_SetDataMaster(FromIS, DValue.CODE, 13));

                          
                        }
                        try
                        {
                            tb_QCHD qh = db.tb_QCHDs.Where(w => w.QCNo.Equals(QCNo1)).FirstOrDefault();
                            if (qh != null)
                            {
                                if (TypeReport.Equals("HINO"))
                                {
                                 //   InsertToExcel(ref worksheet, "G26", db.get_QC_DATAPoint(qh.QCNo, "", 12));
                                  //  InsertToExcel(ref worksheet, "G27", db.get_QC_DATAPoint(qh.QCNo, "", 13));
                                }
                                FormISO = qh.FormISO;
                                LotNo = qh.LotNo;
                                //////////Find UserName////////////
                                var uc = db.tb_QCCheckUsers.Where(u => u.QCNo.Equals(QCNo1)).ToList();
                                foreach (var rd in uc)
                                {
                                    DN = dbShowData.CheckDayN(Convert.ToDateTime(rd.ScanDate));

                                    if (rd.UDesc.Equals("ผู้ตรวจสอบ 1"))
                                        cCheckBy1 = rd.UserName;
                                    if (rd.UDesc.Equals("ผู้ตรวจสอบ 2"))
                                        cCheckBy2 = rd.UserName;
                                    if (rd.UDesc.Equals("ผู้ตรวจสอบ 3"))
                                        cCheckBy3 = rd.UserName;
                                }


                                InsertToExcel(ref worksheet, "J7", cCheckBy1);
                                //InsertToExcel(ref worksheet, "K8", cCheckBy2);
                                InsertToExcel(ref worksheet, "H7", db.QC_GetUserName(qh.ApproveBy));
                                QHNo = qh.QCNo;

                                if (db.QC_CheckNG(QHNo) == "P")
                                {
                                    InsertToExcel(ref worksheet, "J10", "P");
                                }
                                else
                                {
                                    InsertToExcel(ref worksheet, "H10", "P");
                                }

                                //Excel.Range O19 = worksheet.get_Range("O19");
                                //O19.Value2 = db.get_QC_DATAPoint(qh.QCNo, "", 8);
                                //Excel.Range O20 = worksheet.get_Range("O20");
                                //O20.Value2 = db.get_QC_DATAPoint(qh.QCNo, "", 9);
                            }

                        }
                        catch { }

                    }

                    ////////////////////////////////////////
                    ////Insert Desc.//
                    int countA = 0;
                    //row1 = 14;
                    //string setPart = "";
                    //var vList = db.tb_QCGroupParts.Where(p => p.FormISO.Equals(FromIS) && p.PartNo.Equals(PartNo) && !p.TopPic.Trim().Equals("")
                    //            && p.GroupPart.Trim().Equals("")).OrderBy(s => s.Seq).ToList();
                    //foreach (var ls in vList)
                    //{
                    //    countA += 1;
                    //    row1 += 1;
                    //    setPart = ls.SetData.Trim();
                    //    if (ls.TopPic.Trim().Equals("Lot No."))
                    //    {
                    //        setPart = LotRemark;
                    //    }
                    //    InsertToExcel(ref worksheet, "A" + row1.ToString(), countA);
                    //    InsertToExcel(ref worksheet, "B" + row1.ToString(), ls.TopPic.Trim());
                    //    InsertToExcel(ref worksheet, "D" + row1.ToString(), setPart);

                    //}

                    ////////////////


                    countA = 0;
                    string Colm = "";             
                    var listPoint = db.sp_46_QCSelectWO_09_QCTAGSelectPD(QHNo).ToList();
                    if (listPoint.Count > 0)
                    {
                        foreach (var rs in listPoint)
                        {
                            countA += 1;
                            if (true)
                            {
                                row1 = 14;
                                var listPart = db.tb_QCGroupParts.Where(p => p.FormISO.Equals(FromIS) && p.PartNo.Equals(PartNo) && !p.TopPic.Trim().Equals("")
                                                && p.GroupPart.Trim().Equals("")).OrderBy(s => s.Seq).ToList();
                                foreach (var rd in listPart)
                                {
                                    row1 += 1;
                                    if (!rd.SetData.Equals("") && row1 <= 40)
                                    {
                                        try
                                        {

                                            var gValue = db.sp_46_QCGetValue5601(rs.BarcodeTag, QHNo, rd.Seq).FirstOrDefault();
                                            PV = "OK";

                                            if (gValue.CountA > 0)
                                            {
                                                PV = "OK";
                                                if (gValue.CountA > 0)
                                                    PV = "NG";
                                                if (gValue.CountA == 99)
                                                    PV = "";
                                            }



                                            if (rs.BarcodeTag.Contains("Head"))
                                            {
                                                Colm = "H";
                                            }
                                            else if (rs.BarcodeTag.Contains("Middle"))
                                            {
                                                Colm = "J";
                                            }
                                            else
                                            {
                                                Colm = "K";
                                            }                                          
                                            if (row1 == 18)
                                            {
                                                PV = db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, rd.Seq);
                                            }
                                                //Input Value///
                                                InsertToExcel(ref worksheet, Colm + row1.ToString(), PV);
                                            if (TypeReport.Equals("HINO"))
                                            {
                                                if(rd.Seq==12)
                                                {
                                                    InsertToExcel(ref worksheet, Colm + row1.ToString(), db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, 12));
                                                }
                                                if(rd.Seq==13)
                                                {
                                                    InsertToExcel(ref worksheet, Colm + row1.ToString(), db.get_QC_DATAPoint(QHNo, rs.BarcodeTag, 13));
                                                }                                               
                                                //   InsertToExcel(ref worksheet, "G26", db.get_QC_DATAPoint(qh.QCNo, "", 12));
                                                //  InsertToExcel(ref worksheet, "G27", db.get_QC_DATAPoint(qh.QCNo, "", 13));
                                            }
                                        }
                                        catch (Exception ex) { MessageBox.Show(ex.Message); }
                                    }

                                }//foreach
                            }//cunt A
                        }//for
                    }



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
        }

        public static string getDrawing(string part)
        {
            string DW = "";
            using (DataClasses1DataContext db = new DataClasses1DataContext())
            {
                DW = db.QC_GetDrawing(part);
            }
            return DW;

        }

        /// <summary>
        /// //////////////////////////////////////////////////////
        /// </summary>
        public static string CheckColumnTime(string Timex)
        {
            //cboTime.Items.Clear();
            //cboTime.Items.Add("11:00-12:00");
            //cboTime.Items.Add("12:00-13:00");
            //cboTime.Items.Add("17:00-17:30");
            //cboTime.Items.Add("17:30-18:00");

            //cboTime.Items.Add("23:00-24:00");
            //cboTime.Items.Add("00:00-01:00");
            //cboTime.Items.Add("05:00-05:30");
            //cboTime.Items.Add("05:30-06:00");
            string Time = "C";
            if (Timex.Equals("11:00-12:00"))
                Time = "C";
            else if (Timex.Equals("12:00-13:00"))
                Time = "F";
            else if (Timex.Equals("17:00-17:30"))
                Time = "I";
            else if (Timex.Equals("17:30-18:00"))
                Time = "L";
            else if (Timex.Equals("23:00-24:00"))
                Time = "O";
            else if (Timex.Equals("00:00-01:00"))
                Time = "R";
            else if (Timex.Equals("05:00-05:30"))
                Time = "T";
            else if (Timex.Equals("05:30-06:00"))
                Time = "W";

            return Time;
        }
        public static string GetReportName(string defaultx, string PartNo, string FromIS)
        {
            string RP = defaultx;
            using (DataClasses1DataContext db = new DataClasses1DataContext())
            {
                tb_QCGroupPart gp = db.tb_QCGroupParts.Where(p => p.PartNo.Equals(PartNo) && p.FormISO.Equals(FromIS) && p.StepPart.Equals("USE REPORT")
                && p.TopPic.Equals("Report")).FirstOrDefault();
                if (gp != null)
                {
                    
                        RP = gp.SetData.ToString();
                    
                }

            }

            return RP;
        }
        public static string GetSetDataDWG(string PartNo, string FromIS)
        {
            string RP = "";
            using (DataClasses1DataContext db = new DataClasses1DataContext())
            {
                tb_QCGroupPart gp = db.tb_QCGroupParts.Where(p => p.PartNo.Equals(PartNo) && p.FormISO.Equals(FromIS) && p.StepPart.Trim().ToUpper().Equals("DWG  NO.")
                && p.GroupPart.ToUpper().Equals("YES")).FirstOrDefault();
                if (gp != null)
                {
                    RP = gp.SetData.ToString();
                }

            }

            return RP;
        }
        public static string GetSetDataCust(string PartNo, string FromIS)
        {
            string RP = "";
            using (DataClasses1DataContext db = new DataClasses1DataContext())
            {
                tb_QCGroupPart gp = db.tb_QCGroupParts.Where(p => p.PartNo.Equals(PartNo) && p.FormISO.Equals(FromIS) && p.StepPart.Trim().ToUpper().Equals("CUSTOMER REF.")
                && p.GroupPart.ToUpper().Equals("YES")).FirstOrDefault();
                if (gp != null)
                {
                    RP = gp.SetData.ToString();
                }

            }

            return RP;
        }
        public static string GetSetDataLB(string PartNo, string FromIS,string Topic)
        {
            string RP = "";
            using (DataClasses1DataContext db = new DataClasses1DataContext())
            {
                tb_QCGroupPart gp = db.tb_QCGroupParts.Where(p => p.PartNo.Equals(PartNo) && p.FormISO.Equals(FromIS) && p.StepPart.Trim().ToUpper().Equals(Topic.ToUpper())).FirstOrDefault();
                if (gp != null)
                {
                    RP = gp.SetData.ToString();
                }

            }

            return RP;
        }
        public static string GetSetData(string PartNo, string FromIS,int Seq)
        {
            string RP = "";
            using (DataClasses1DataContext db = new DataClasses1DataContext())
            {
                tb_QCGroupPart gp = db.tb_QCGroupParts.Where(p => p.PartNo.Equals(PartNo) && p.FormISO.Equals(FromIS) && p.Seq.Equals(Seq)).FirstOrDefault();
                if (gp != null)
                {
                    RP = gp.SetData.ToString();
                }

            }

            return RP;
        }

        public static void InsertToExcel(ref Excel.Worksheet exc,string Column,dynamic Values)
        {
            try
            {                
                Excel.Range refs = exc.get_Range(Column);
                refs.Value2 = Values;
                if (Values.Contains("Æ"))
                {
                    int addint = Values.IndexOf("Æ");
                    refs.Characters[addint, 2].Font.Name = "Symbol";//                          
                }
            }
            catch { }
        }
        public static void InsertQCChecker(string Uid,string QCNo,string TypeS,string Desc,string BoxNo)
        {
            try
            {
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    if (!Desc.Equals(""))
                    {
                        string DN= dbShowData.CheckDayN(DateTime.Now);
                        tb_QCCheckUser cu = db.tb_QCCheckUsers.Where(u => u.UserID.Equals(Uid) && u.QCNo.Equals(QCNo) && u.UType.Equals(TypeS)
                        && u.UDesc.Equals(Desc)
                        && u.DayN.Equals(DN)
                        ).FirstOrDefault();
                        if (cu != null)
                        {

                        }
                        else
                        {
                            tb_User us = db.tb_Users.Where(u => u.UserID.Equals(Uid)).FirstOrDefault();
                            if (us != null)
                            {

                                tb_QCCheckUser cn = new tb_QCCheckUser();
                                cn.QCNo = QCNo;
                                cn.ScanDate = DateTime.Now;
                                cn.UserID = Uid;
                                cn.UserName = us.NameApp;
                                cn.UType = TypeS;
                                cn.DayN = DN;
                                cn.UDesc = Desc;
                                cn.BoxNo = BoxNo;
                                db.tb_QCCheckUsers.InsertOnSubmit(cn);
                                db.SubmitChanges();
                            }
                        }
                    }
                }
            }
            catch { }
        }
        public static void DeleteQCChecker(int id)
        {
            try
            {
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    tb_QCCheckUser cx = db.tb_QCCheckUsers.Where(w => w.id.Equals(id)).FirstOrDefault();
                    if (cx != null)
                    {
                        db.tb_QCCheckUsers.DeleteOnSubmit(cx);
                        db.SubmitChanges();
                    }
                }

            }
            catch { }
        }

        private static string[] ConvertToStringArray(System.Array values)
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
        private static void releaseObject(object obj)
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
        private static string Getcolumn(int Col)
        {
            string RT = "";

            if (Col.Equals(1))
                RT = "A";
            else if (Col.Equals(2))
                RT = "B";
            else if (Col.Equals(3))
                RT = "C";
            else if (Col.Equals(4))
                RT = "D";
            else if (Col.Equals(5))
                RT = "E";
            else if (Col.Equals(6))
                RT = "F";
            else if (Col.Equals(7))
                RT = "G";
            else if (Col.Equals(8))
                RT = "H";
            else if (Col.Equals(9))
                RT = "I";
            else if (Col.Equals(10))
                RT = "J";
            else if (Col.Equals(11))
                RT = "K";
            else if (Col.Equals(12))
                RT = "L";
            else if (Col.Equals(13))
                RT = "M";
            else if (Col.Equals(14))
                RT = "N";
            else if (Col.Equals(15))
                RT = "O";
            else if (Col.Equals(16))
                RT = "P";
            else if (Col.Equals(17))
                RT = "Q";
            else if (Col.Equals(18))
                RT = "R";
            else if (Col.Equals(19))
                RT = "S";
            else if (Col.Equals(20))
                RT = "T";
            else if (Col.Equals(21))
                RT = "U";
            else if (Col.Equals(22))
                RT = "V";
            else if (Col.Equals(23))
                RT = "W";
            else if (Col.Equals(24))
                RT = "X";
            else if (Col.Equals(25))
                RT = "Y";
            else if (Col.Equals(26))
                RT = "Z";

           else if (Col.Equals(27))
                RT = "AA";
            else if (Col.Equals(28))
                RT = "AB";
            else if (Col.Equals(29))
                RT = "AC";
            else if (Col.Equals(30))
                RT = "AD";
            else if (Col.Equals(31))
                RT = "AE";
            else if (Col.Equals(32))
                RT = "AF";
            else if (Col.Equals(33))
                RT = "AG";
            else if (Col.Equals(34))
                RT = "AH";
            else if (Col.Equals(35))
                RT = "AI";
            else if (Col.Equals(36))
                RT = "AJ";
            else if (Col.Equals(37))
                RT = "AK";
            else if (Col.Equals(38))
                RT = "AL";
            else if (Col.Equals(39))
                RT = "AM";
            else if (Col.Equals(40))
                RT = "AN";
            else if (Col.Equals(41))
                RT = "AO";
            else if (Col.Equals(42))
                RT = "AP";
            else if (Col.Equals(43))
                RT = "AQ";

            else if (Col.Equals(44))
                RT = "AR";
            else if (Col.Equals(45))
                RT = "AS";
            else if (Col.Equals(46))
                RT = "AT";
            else if (Col.Equals(47))
                RT = "AU";
            else if (Col.Equals(48))
                RT = "AV";
            else if (Col.Equals(49))
                RT = "AW";
            else if (Col.Equals(50))
                RT = "AX";


            return RT;
        }
        private static string GetDWG(string PartNo,int AC)
        {
            string RT = "";
            using (DataClasses1DataContext db3 = new DataClasses1DataContext())
            {
                RT = db3.getItemNo_DWG(PartNo, AC);
            }

            return RT;
        }
    }
}
