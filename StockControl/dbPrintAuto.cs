using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using CrystalDecisions.CrystalReports.Engine;

namespace StockControl
{
    public static class dbPrintAuto
    {
        public static void PrintTAGA()
        {
            bool printA = false;
            
            using (DataClasses1DataContext db = new DataClasses1DataContext())
            {

                var pa = db.tb_PrintAutos.Where(p=>p.PrintFlag.Equals(false) && p.PrintApp.Equals("Export")).ToList();
                foreach(var rd in pa)
                {
                    //Print A
                    tb_ExportDetail exp = db.tb_ExportDetails.Where(b => b.id.Equals(Convert.ToInt32(rd.Refid))).FirstOrDefault();
                    if (exp!=null)
                    {
                        if (exp.PrintType.Equals("A"))
                        {
                            //Call
                            printA = false;
                            printA = printBigA(Convert.ToInt32(rd.Refid), rd.PrintDocuNo);
                            //end Call
                            tb_PrintAuto EPA = db.tb_PrintAutos.Where(i => i.id.Equals(rd.id)).FirstOrDefault();
                            if (EPA != null && printA)
                            {
                                EPA.PrintFlag = true;
                                EPA.PrintDate = DateTime.Now;
                                db.SubmitChanges();
                            }
                        }
                        else
                        {

                        }
                    }
                    
                }

            }
        }
        public static void PrintTAGB()
        {
            return;
            bool printA = false;

            using (DataClasses1DataContext db = new DataClasses1DataContext())
            {

                var pa = db.tb_PrintAutos.Where(p => p.PrintFlag.Equals(false) && p.PrintApp.Equals("Export")).ToList();
                foreach (var rd in pa)
                {
                    //Print A
                    tb_ExportDetail exp = db.tb_ExportDetails.Where(b => b.id.Equals(Convert.ToInt32(rd.Refid))).FirstOrDefault();
                    if (exp != null)
                    {
                        if (exp.PrintType.Equals("B"))
                        {
                            //Call
                            printA = false;
                            printA = printTypeB(Convert.ToInt32(rd.Refid), rd.PrintDocuNo);
                            //end Call
                            tb_PrintAuto EPA = db.tb_PrintAutos.Where(i => i.id.Equals(rd.id)).FirstOrDefault();
                            if (EPA != null && printA)
                            {
                                EPA.PrintFlag = true;
                                EPA.PrintDate = DateTime.Now;
                                db.SubmitChanges();
                            }
                        }
                        else
                        {

                        }
                    }

                }

            }
        }
        private static bool printBigA(int id,string ExNo)
        {
            bool RT = false;
            try
            {
                string DATA = "";
                
                DATA = AppDomain.CurrentDomain.BaseDirectory;
                DATA = DATA + @"Report\" + "ExInvoiceTAX.rpt";
                //radGridView1.EndEdit();
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    db.sp_014_CreateGroup_Dynamics(ExNo);
                    db.sp_014_DeletePrintTAG(ExNo);
                    string QRCode = "";
                    string Country = "";
                    string CountrySize = "";
                    int TotalPallet = 1;
                    int plNo = 0;
                    //TotalPallet=Convert.ToInt32(txtTotalPallet.Text);

                    tb_ExportList el = db.tb_ExportLists.Where(w => w.InvoiceNo == ExNo).FirstOrDefault();
                    if (el != null)
                    {
                        CountrySize = el.CountrySize;
                        Country = el.Country;
                    }

                    var exDetail = db.tb_ExportDetails.Where(e => e.InvoiceNo.Equals(ExNo)).ToList();
                    foreach(var rs in exDetail)
                    {
                        plNo = 0;
                        int.TryParse(rs.PalletNo,out plNo);
                        if (plNo>TotalPallet)
                        {
                            TotalPallet = plNo;
                        }
                    }

                    //foreach (GridViewRowInfo rd in radGridView1.Rows)
                    //{

                        if (true)
                        {

                            tb_ExportDetail ed = db.tb_ExportDetails.Where(ee => ee.id == Convert.ToInt32(id) && ee.PrintType == "A").FirstOrDefault();
                            if (ed != null)
                            {
                                DateTime SH = Convert.ToDateTime(ed.ShippingDate);
                                tb_ExportList exx = db.tb_ExportLists.Where(es => es.InvoiceNo == ed.InvoiceNo).FirstOrDefault();
                                if (exx != null)
                                {
                                    SH = Convert.ToDateTime(exx.LoadDate);
                                }
                                string getNewLot = ed.LotNo;
                                if (getNewLot.Equals(""))
                                {
                                    //PD,WO22104844,4,375,23VT,51of94,44130036090,290322                                   
                                    string[] GT = Convert.ToString(db.getLOtExport(ed.PartNo, ed.OrderNo, ed.InvoiceNo, ed.id)).Split(',');
                                    if (GT.Length > 4)
                                    {
                                        getNewLot = GT[4];
                                    }
                                    if (!getNewLot.Equals(""))
                                    {
                                        db.sp_019_LocaDeliveryList_DynamicsUpdateLot(ed.id, getNewLot);
                                    }
                                }

                                //Order,PalletNo,Invoice,PartCode,Qty,ofTAG,TotalTAG,LotNo
                                QRCode = "";
                                QRCode = ed.OrderNo + "," + ed.PalletNo + "," + ed.InvoiceNo + ",";
                                QRCode = QRCode + ed.PartNo + "," + ed.Qty + "," + ed.ofPL.ToString() + "of" + ed.TotalPL.ToString() + "," + getNewLot.ToString();
                                byte[] barcode = dbClss.SaveQRCode2D(QRCode);

                                tb_ExportPrintTAG ep = new tb_ExportPrintTAG();
                                ep.CustomerAddress = "5-1 Kanaya, Murayama, Yamagata, 995-0004 Japan";// ed.CustomerAddress.ToString();
                                //MessageBox.Show("OK");
                                ep.CustomerItemName = ""; //// ed.CustItem.ToString();                               
                                ep.CustomerItemNo = Convert.ToString(db.getItemCSTM_Dynamics(ed.PartNo, ""));
                                ep.CustomerName = "Nabtesco Autmotive Corporation";// Convert.ToString(db.getItemCSTMName(ed.Customer));

                                if (Country.ToUpper().Equals("INDIA"))
                                {
                                    ep.CustomerName = "MINDA NABTESCO AUTOMOTIVE PVT LTD";
                                    ep.CustomerAddress = "Plot no-191 sector-8 IMT  Manesar ,distt- Gurgaon- 122050 State Haryana.";
                                }
                                ep.InvoiceNo = ExNo;


                                ep.LOTNo = getNewLot;
                                ep.QRCode = barcode;

                                ep.Qty = ed.Qty;
                                ep.GroupP = ed.GroupP;
                                ep.ShippingDate = SH;// ed.ShippingDate;
                                ep.TotalPLOf = Convert.ToInt32(ed.PalletNo);
                                ep.TotalPLOfQty = Convert.ToInt32(TotalPallet);
                                ep.PLOfQty = Convert.ToInt32(ed.TotalPL);
                                ep.PLOf = Convert.ToInt32(ed.ofPL);
                                ep.PartCode = ed.PartNo;
                                ep.PartName = ed.PartName;
                                ep.Country = Country;
                                ep.CountrySize = CountrySize;


                                db.tb_ExportPrintTAGs.InsertOnSubmit(ep);
                                db.SubmitChanges();
                            }

                        }
                    //}
                    try
                    {
                        ReportDocument RPT = new ReportDocument();
                        RPT.Load(DATA);
                        Report.Reportx1.SetDataSourceConnection(RPT);
                        RPT.SetParameterValue("@InvoiceNo", Convert.ToString(ExNo));
                        RPT.SetParameterValue("@Datex", DateTime.Now);
                        RPT.PrintToPrinter(1, false, 1, 1);
                        ////Print//
                        ////
                        RPT.Close();
                        RPT.Dispose();
                        RT= true;
                    }
                    catch { }
                    
                    /*
                    Report.Reportx1.WReport = "PrintEXTAG";
                    Report.Reportx1.Value = new string[1];
                    Report.Reportx1.Value[0] = ExNo;
                    Report.Reportx1 op = new Report.Reportx1("ExInvoiceTAX.rpt");
                    op.Show();
                    */

                }
            }
            catch (Exception ex) { }
            return RT;
        }
        private static bool printTypeB(int id,string ExNo)
        {

            bool RT = false;
            /*
            string DATA = "";
            DATA = AppDomain.CurrentDomain.BaseDirectory;
            DATA = DATA + @"Report\" + "ExInvoiceTAX2.rpt";
            try
            {
                using (DataClasses1DataContext db = new DataClasses1DataContext())
                {
                    db.sp_014_CreateGroup_Dynamics(ExNo);
                    db.sp_014_DeletePrintTAG(ExNo);
                    string QRCode = "";
                    string Country = "";
                    string CountrySize = "";
                    
                    DateTime SH = DateTime.Now;// Convert.ToDateTime(ed.ShippingDate);


                    tb_ExportList el = db.tb_ExportLists.Where(w => w.InvoiceNo == ExNo).FirstOrDefault();
                    if (el != null)
                    {
                        CountrySize = el.CountrySize;
                        Country = el.Country;
                        SH = Convert.ToDateTime(el.LoadDate);
                    }
                    if(true)
                    {
                        if (true)
                        {
                            tb_ExportDetail ed = db.tb_ExportDetails.Where(ee => ee.id == Convert.ToInt32(rd.Cells["id"].Value) && ee.PrintType == "B").FirstOrDefault();
                            if (ed != null)
                            {
                                string getNewLot = ed.LotNo;
                                if (getNewLot.Equals(""))
                                {
                                    //PD,WO22104844,4,375,23VT,51of94,44130036090,290322                                   
                                    string[] GT = Convert.ToString(db.getLOtExport(ed.PartNo, ed.OrderNo, ed.InvoiceNo, ed.id)).Split(',');
                                    if (GT.Length > 4)
                                    {
                                        getNewLot = GT[4];
                                    }
                                    if (!getNewLot.Equals(""))
                                    {
                                        db.sp_019_LocaDeliveryList_DynamicsUpdateLot(ed.id, getNewLot);
                                    }
                                }
                                //Order,PalletNo,Invoice,PartCode,Qty,ofTAG,TotalTAG,LotNo

                                QRCode = "";
                                QRCode = ed.OrderNo + "," + ed.PalletNo + "," + ed.InvoiceNo + "," + ed.PartNo + "," + ed.Qty + "," + ed.ofPL.ToString() + "of" + ed.TotalPL.ToString() + "," + getNewLot.ToString();
                                byte[] barcode = dbClss.SaveQRCode2D(QRCode);
                                tb_ExportPrintTAG ep = new tb_ExportPrintTAG();
                                ep.CustomerAddress = "5-1 Kanaya, Murayama, Yamagata, 995-0004 Japan";
                                //ep.CustomerAddress = ed.CustomerAddress;//"5-1 Kanaya,Murayama,Yamagata,995-0004 Japan.";
                                ep.CustomerItemName = "";// ep.CustomerItemName;
                                ep.CustomerItemNo = Convert.ToString(db.getItemCSTM_Dynamics(ed.PartNo, "")); // ed.CustItem;                             
                                                                                                              // ep.CustomerName = ed.CustomerName;// Convert.ToString(db.getItemCSTMName(ed.Customer));
                                ep.CustomerName = "Nabtesco Autmotive Corporation";// Convert.ToString(db.getItemCSTMName(ed.Customer));
                                ep.InvoiceNo = txtExportNo.Text;

                                ep.LOTNo = getNewLot;
                                ep.QRCode = barcode;
                                ep.Qty = ed.Qty;
                                ep.GroupP = ed.GroupP;
                                ep.ShippingDate = SH;// ed.ShippingDate;
                                ep.TotalPLOf = Convert.ToInt32(ed.PalletNo);
                                ep.TotalPLOfQty = Convert.ToInt32(txtTotalPallet.Text);
                                ep.PLOfQty = Convert.ToInt32(ed.TotalPL);
                                ep.PLOf = Convert.ToInt32(ed.ofPL);
                                ep.PartCode = ed.PartNo;
                                ep.PartName = ed.PartName;
                                ep.Country = Country;
                                ep.CountrySize = CountrySize;
                                if (Country.ToUpper().Equals("INDIA"))
                                {
                                    ep.CustomerName = "MINDA NABTESCO AUTOMOTIVE PVT LTD";
                                    ep.CustomerAddress = "Plot no-191 sector-8 IMT  Manesar ,distt- Gurgaon- 122050 State Haryana.";
                                }

                                db.tb_ExportPrintTAGs.InsertOnSubmit(ep);
                                db.SubmitChanges();
                            }

                        }
                    }

                    Report.Reportx1.WReport = "PrintEXTAG";
                    Report.Reportx1.Value = new string[1];
                    Report.Reportx1.Value[0] = ExNo;
                    Report.Reportx1 op = new Report.Reportx1("ExInvoiceTAX2.rpt");
                    op.Show();
                }
            }
            catch { }
            */

            return RT;
        }
    }
}
