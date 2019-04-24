using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using NPOI.XSSF.UserModel;
using System.IO;

/// <summary>
///ExcelOperation 的摘要说明
/// </summary>
public class ExcelOperation
{
	public ExcelOperation()
	{
		//Get coloum header from Template.xls
        #region Coloum header
        excelHead.Add(new KeyValuePair<int, string>(0, "SOrg."));
        excelHead.Add(new KeyValuePair<int, string>(1, "SaTy"));
        excelHead.Add(new KeyValuePair<int, string>(2, "Sales Doc."));
        excelHead.Add(new KeyValuePair<int, string>(3, "SalesItem"));
        excelHead.Add(new KeyValuePair<int, string>(4, "DLV Plant"));
        excelHead.Add(new KeyValuePair<int, string>(5, "Prod Ord Type"));
        excelHead.Add(new KeyValuePair<int, string>(6, "Production Order"));
        excelHead.Add(new KeyValuePair<int, string>(7, "Status"));
        excelHead.Add(new KeyValuePair<int, string>(8, "Material"));
        excelHead.Add(new KeyValuePair<int, string>(9, "Sold-To Pt"));
        excelHead.Add(new KeyValuePair<int, string>(10, "Sold-To Party"));
        excelHead.Add(new KeyValuePair<int, string>(11, "Segment"));
        excelHead.Add(new KeyValuePair<int, string>(12, "Personnel"));
        excelHead.Add(new KeyValuePair<int, string>(13, "Sales Responsible"));
        excelHead.Add(new KeyValuePair<int, string>(14, "SOff."));
        excelHead.Add(new KeyValuePair<int, string>(15, "Sales Office"));
        excelHead.Add(new KeyValuePair<int, string>(16, "Product Hierarchy"));
        excelHead.Add(new KeyValuePair<int, string>(17, "WBS Element"));
        excelHead.Add(new KeyValuePair<int, string>(18, "WBS Element"));
        excelHead.Add(new KeyValuePair<int, string>(19, "Project Description"));
        excelHead.Add(new KeyValuePair<int, string>(20, "Proj Engineer"));
        excelHead.Add(new KeyValuePair<int, string>(21, "Proj Engineer Name"));
        excelHead.Add(new KeyValuePair<int, string>(22, "Net value"));
        excelHead.Add(new KeyValuePair<int, string>(23, "Order Quantity"));
        excelHead.Add(new KeyValuePair<int, string>(24, "Confirmed Qty"));
        excelHead.Add(new KeyValuePair<int, string>(25, "SO CreatedOn"));
        excelHead.Add(new KeyValuePair<int, string>(26, "Prod CreatedOn"));
        excelHead.Add(new KeyValuePair<int, string>(27, "Actual Release Date"));
        excelHead.Add(new KeyValuePair<int, string>(28, "Actual Finish Date"));
        excelHead.Add(new KeyValuePair<int, string>(29, "1st GR for Prod Order"));
        excelHead.Add(new KeyValuePair<int, string>(30, "Quotation LT"));
        excelHead.Add(new KeyValuePair<int, string>(31, "1st ConfirmedDt"));
        excelHead.Add(new KeyValuePair<int, string>(32, "Requested Date"));
        excelHead.Add(new KeyValuePair<int, string>(33, "Basic End Date"));
        excelHead.Add(new KeyValuePair<int, string>(34, "ShipmentStartOn"));
        excelHead.Add(new KeyValuePair<int, string>(35, "Route"));
        excelHead.Add(new KeyValuePair<int, string>(36, "Route"));
        excelHead.Add(new KeyValuePair<int, string>(37, "TransitTime"));
        excelHead.Add(new KeyValuePair<int, string>(38, "Delivery Number"));
        excelHead.Add(new KeyValuePair<int, string>(39, "Shipment Number"));

        excelHeader[0] = "SOrg.";
        excelHeader[1] = "SaTy";
        excelHeader[2] = "Sales Doc.";
        excelHeader[3] = "SalesItem";
        excelHeader[4] = "DLV Plant";
        excelHeader[5] = "Prod Ord Type";
        excelHeader[6] = "Production Order";
        excelHeader[7] = "Status";
        excelHeader[8] = "Material";
        excelHeader[9] = "Sold-To Pt";
        excelHeader[10] = "Sold-To Party";
        excelHeader[11] = "Segment";
        excelHeader[12] = "Personnel";
        excelHeader[13] = "Sales Responsible";
        excelHeader[14] = "SOff.";
        excelHeader[15] = "Sales Office";
        excelHeader[16] = "Product Hierarchy";
        excelHeader[17] = "WBS Element";
        excelHeader[18] = "WBS Element";
        excelHeader[19] = "Project Description";
        excelHeader[20] = "Proj Engineer";
        excelHeader[21] = "Proj Engineer Name";
        excelHeader[22] = "Net value";
        excelHeader[23] = "Order Quantity";
        excelHeader[24] = "Confirmed Qty";
        excelHeader[25] = "SO CreatedOn";
        excelHeader[26] = "Prod CreatedOn";
        excelHeader[27] = "Actual Release Date";
        excelHeader[28] = "Actual Finish Date";
        excelHeader[29] = "1st GR for Prod Order";
        excelHeader[30] = "Quotation LT";
        excelHeader[31] = "1st ConfirmedDt";
        excelHeader[32] = "Requested Date";
        excelHeader[33] = "Basic End Date";
        excelHeader[34] = "ShipmentStartOn";
        excelHeader[35] = "Route";
        excelHeader[36] = "Route";
        excelHeader[37] = "TransitTime";
        excelHeader[38] = "Delivery Number";
        excelHeader[39] = "Shipment Number"; 
        #endregion
	}
    public string[] excelHeader = new string[40];
    public List<KeyValuePair<int, string>> excelHead = new List<KeyValuePair<int, string>>();
    public XSSFWorkbook wb;
    private XSSFSheet sht;
    public int rowsCount, colsCount;

    public bool ExcelImportWithLayoutCheck(string filename, string sheetName = "Sheet1")
    {
        bool bool_ImportResult = true;
        wb = new XSSFWorkbook(File.OpenRead(@"C:\Users\Public\Music\" + filename));
        //wb = new XSSFWorkbook(File.OpenRead(@"F:\" + filename));
        sht = (XSSFSheet)wb.GetSheet(sheetName);
        if (sht==null)
        {
            sht = (XSSFSheet)wb.GetSheetAt(0);
        }
        //取行Excel的最大行数
        rowsCount = sht.PhysicalNumberOfRows;
        //取所有行中的最大列数
        colsCount = sht.GetRow(0).PhysicalNumberOfCells;
        for (int i = 0; i < excelHeader.Length; i++)
        {
            if (((XSSFCell)((XSSFRow)sht.GetRow(0)).GetCell(i)).StringCellValue != excelHeader[i])
            { 
                bool_ImportResult = false;
                break;
            }
        }
        return bool_ImportResult;
    }

    public void TraceFromExcel()
    {
        //Clean last count and array init
        Data2Trace.CleanCounter();
        for (int i = 0; i < mylist_SO_perMonth.Length; i++)
        {
            mylist_SO_perMonth[i] = new List<string>();
        }
        for (int i = 0; i < mylist_0400_LT_perMonth.Length; i++)
        {
            mylist_0400_LT_perMonth[i] = new List<int>();
        }
        for (int i = 0; i < mylist_0400_LT_perWeek.Length; i++)
        {
            mylist_0400_LT_perWeek[i] = new List<int>();
        }
        for (int i = 0; i < mylist_0481_LT_perMonth.Length; i++)
        {
            mylist_0481_LT_perMonth[i] = new List<int>();
        }
        for (int i = 0; i < mylist_0481_LT_perWeek.Length; i++)
        {
            mylist_0481_LT_perWeek[i] = new List<int>();
        }
        for (int i = 0; i < mylist_LT_perMonth.Length; i++)
        {
            mylist_LT_perMonth[i] = new List<int>();
        }
        for (int i = 0; i < mylist_LT_perWeek.Length; i++)
        {
            mylist_LT_perWeek[i] = new List<int>();
        }
        //逐行统计
        for (int i = 1; i < rowsCount; i++)
        {
            RowRecordProcess(i);
        }
        //ECharts5&ECharts6需要处理一下原始数据
        //ECharts5 SO Create perMonth
        for (int i = 0; i < mylist_SO_perMonth.Length; i++)
        {
            List<string> nlist = mylist_SO_perMonth[i].Distinct().ToList();
            Data2Trace.count_SO_soCreated_perMonth[i] = nlist.Count();
            Data2Trace.count_SO_soCreated_YTD = Data2Trace.count_SO_soCreated_YTD + Data2Trace.count_SO_soCreated_perMonth[i];
        }
        //ECharts6 LT(Average Lead Time：CDS)平均交货时间        
        for (int i = 0; i < mylist_0400_LT_perMonth.Length; i++)
        {
            if (mylist_0400_LT_perMonth[i].Count == 0)
            {
                Data2Trace.LT_0400_perMonth[i] = 0;

            }
            else
            {
                Data2Trace.LT_0400_perMonth[i] = mylist_0400_LT_perMonth[i].Average();

            }
        }
        if (mylist_0400_LT_YTD.Count == 0)
        {
            Data2Trace.LT_0400_YTD = 0;

        }
        else
        {
            Data2Trace.LT_0400_YTD = mylist_0400_LT_YTD.Average();

        }

        for (int i = 0; i < mylist_0481_LT_perMonth.Length; i++)
        {
            if (mylist_0481_LT_perMonth[i].Count == 0)
            {
                Data2Trace.LT_0481_perMonth[i] = 0;

            }
            else
            {
                Data2Trace.LT_0481_perMonth[i] = mylist_0481_LT_perMonth[i].Average();

            }
        }
        if (mylist_0481_LT_YTD.Count == 0)
        {
            Data2Trace.LT_0481_YTD = 0;

        }
        else
        {
            Data2Trace.LT_0481_YTD = mylist_0481_LT_YTD.Average();

        }

        for (int i = 0; i < mylist_LT_perMonth.Length; i++)
        {
            if (mylist_LT_perMonth[i].Count == 0)
            {
                Data2Trace.LT_perMonth[i] = 0;

            }
            else
            {
                Data2Trace.LT_perMonth[i] = mylist_LT_perMonth[i].Average();

            }
        }
        if (mylist_LT_YTD.Count == 0)
        {
            Data2Trace.LT_YTD = 0;

        }
        else
        {
            Data2Trace.LT_YTD = mylist_LT_YTD.Average();

        }
    }

    List<string>[] mylist_SO_perMonth = new List<string>[12];
    List<string> mylist_SO_YTD = new List<string>();
    List<int>[] mylist_0400_LT_perMonth = new List<int>[12];
    List<int>[] mylist_0400_LT_perWeek = new List<int>[4];
    List<int> mylist_0400_LT_YTD = new List<int>();
    List<int>[] mylist_0481_LT_perMonth = new List<int>[12];
    List<int>[] mylist_0481_LT_perWeek = new List<int>[4];
    List<int> mylist_0481_LT_YTD = new List<int>();
    List<int>[] mylist_LT_perMonth = new List<int>[12];
    List<int>[] mylist_LT_perWeek = new List<int>[4];
    List<int> mylist_LT_YTD = new List<int>();

    public bool RowRecordProcess(int rowIndex)
    {
        try
        {
            //提取关注的信息到监控类中并统计部分数据
            XSSFRow nRow = (XSSFRow)sht.GetRow(rowIndex);
            if (((XSSFCell)nRow.GetCell(2)).CellType == NPOI.SS.UserModel.CellType.Blank)
            {
                //销售订单号为空，以此判断此行为空
            }
            else
            {
                #region 该行有数据
                bool bool_noLT = false;
                bool bool_noBasicEndDate = false;
                #region 赋值
                try
                {
                    #region SO,每单都应该有
                    if (((XSSFCell)nRow.GetCell(2)).CellType == NPOI.SS.UserModel.CellType.String)
                    {
                        ProcessMonitor.mSO = ((XSSFCell)nRow.GetCell(2)).StringCellValue;//每单都有

                    }
                    else if (((XSSFCell)nRow.GetCell(2)).CellType == NPOI.SS.UserModel.CellType.Numeric)
                    {
                        ProcessMonitor.mSO = ((XSSFCell)nRow.GetCell(2)).NumericCellValue.ToString();//每单都有

                    }
                    #endregion
                    #region NetValue,每单都有
                    ProcessMonitor.mNetValue = ((XSSFCell)nRow.GetCell(22)).NumericCellValue;                    
                    #endregion
                    #region PO,未创建生产单号的此项为空
                    if (((XSSFCell)nRow.GetCell(6)).CellType == NPOI.SS.UserModel.CellType.String)
                    {
                        ProcessMonitor.mPO = ((XSSFCell)nRow.GetCell(6)).StringCellValue;//未创建生产单号的此项为空

                    }
                    else if (((XSSFCell)nRow.GetCell(6)).CellType == NPOI.SS.UserModel.CellType.Numeric)
                    {
                        ProcessMonitor.mPO = ((XSSFCell)nRow.GetCell(6)).NumericCellValue.ToString();//未创建生产单号的此项为空

                    }
                    #endregion
                    #region DLV_Plant,每单都应该有
                    string str_DLV_Plant = "";
                    if (((XSSFCell)nRow.GetCell(4)).CellType == NPOI.SS.UserModel.CellType.String)
                    {
                        str_DLV_Plant = ((XSSFCell)nRow.GetCell(4)).StringCellValue;

                    }
                    else if (((XSSFCell)nRow.GetCell(4)).CellType == NPOI.SS.UserModel.CellType.Numeric)
                    {
                        str_DLV_Plant = ((XSSFCell)nRow.GetCell(4)).NumericCellValue.ToString();

                    }
                    if (str_DLV_Plant == "0400")
                    {
                        ProcessMonitor.mDLV_Plant = DLV_Plant.P_0400;
                    }
                    else if (str_DLV_Plant == "0481")
                    {
                        ProcessMonitor.mDLV_Plant = DLV_Plant.P_0481;
                    }
                    else
                    {
                        Console.WriteLine("No DLV_Plant");
                        return false;
                    }
                    #endregion
                    //每个PO都应完全的时间节点
                    #region SO_CreatedOn/RequestDate/QuotationLT
                    ProcessMonitor.mActualTime.Date_SO_CreatedOn = DateTime.FromOADate(((XSSFCell)nRow.GetCell(25)).NumericCellValue);
                    ProcessMonitor.mNetValue = ((XSSFCell)nRow.GetCell(22)).NumericCellValue;
                    if (((XSSFCell)nRow.GetCell(32)).CellType == NPOI.SS.UserModel.CellType.Blank)
                    {
                        Console.Write("Without RequestDate");
                    }
                    else
                    {
                        ProcessMonitor.mEstimatedTime.Date_RequestDate = DateTime.FromOADate(((XSSFCell)nRow.GetCell(32)).NumericCellValue);

                    }
                    if (((XSSFCell)nRow.GetCell(30)).CellType == NPOI.SS.UserModel.CellType.Blank)
                    {
                        bool_noLT = true;
                        Console.Write("Without QuotationLT");
                    }
                    else
                    {
                        ProcessMonitor.mEstimatedTime.QuotationLT = (int)((XSSFCell)nRow.GetCell(30)).NumericCellValue;
                        bool_noLT = false;
                    }
                    #endregion
                    //可能没有的时间节点
                    #region 1stConfirmedDt/BasicEndDate
                    if (((XSSFCell)nRow.GetCell(31)).CellType == NPOI.SS.UserModel.CellType.Blank)
                    {
                        //no Date_1stConfirmedDt                    
                    }
                    else
                    {
                        ProcessMonitor.mEstimatedTime.Date_1stConfirmedDt = DateTime.FromOADate(((XSSFCell)nRow.GetCell(31)).NumericCellValue);

                    }

                    if (((XSSFCell)nRow.GetCell(33)).CellType == NPOI.SS.UserModel.CellType.Blank)
                    {
                        bool_noBasicEndDate = true;
                        //no Date_BasicEndDate                   
                    }
                    else
                    {
                        ProcessMonitor.mEstimatedTime.Date_BasicEndDate = DateTime.FromOADate(((XSSFCell)nRow.GetCell(33)).NumericCellValue);
                        bool_noBasicEndDate = false;
                    }
                    #endregion
                    //根据节点有无判断项目进展阶段
                    #region 时间节点赋值，判断项目进展阶段
                    if (((XSSFCell)nRow.GetCell(26)).CellType == NPOI.SS.UserModel.CellType.Blank)
                    {
                        //PO not Created
                        ProcessMonitor.mPO_Status = PO_Status.SO_Created;
                    }
                    else
                    {
                        ProcessMonitor.mActualTime.Date_PO_CreatedOn = DateTime.FromOADate(((XSSFCell)nRow.GetCell(26)).NumericCellValue);
                        if (((XSSFCell)nRow.GetCell(27)).CellType == NPOI.SS.UserModel.CellType.Blank)
                        {
                            //PO not Release
                            ProcessMonitor.mPO_Status = PO_Status.PO_Created;
                        }
                        else
                        {
                            ProcessMonitor.mActualTime.Date_ActualReleaseDate = DateTime.FromOADate(((XSSFCell)nRow.GetCell(27)).NumericCellValue);
                            if (((XSSFCell)nRow.GetCell(28)).CellType == NPOI.SS.UserModel.CellType.Blank)
                            {
                                //PO not Finished
                                ProcessMonitor.mPO_Status = PO_Status.PO_Released;
                            }
                            else
                            {
                                ProcessMonitor.mActualTime.Date_ActualFinishDate = DateTime.FromOADate(((XSSFCell)nRow.GetCell(28)).NumericCellValue);
                                ProcessMonitor.mActualTime.Date_1stGRforPO = DateTime.FromOADate(((XSSFCell)nRow.GetCell(29)).NumericCellValue);
                                if (((XSSFCell)nRow.GetCell(34)).CellType == NPOI.SS.UserModel.CellType.Blank)
                                {
                                    //PO not Shipment
                                    ProcessMonitor.mPO_Status = PO_Status.PO_Finished;
                                }
                                else
                                {
                                    ProcessMonitor.mActualTime.Date_ShipmentStartOn = DateTime.FromOADate(((XSSFCell)nRow.GetCell(34)).NumericCellValue);
                                    ProcessMonitor.mPO_Status = PO_Status.DLV;
                                    if (((XSSFCell)nRow.GetCell(7)).StringCellValue != "DLV")
                                    {
                                        Console.WriteLine("Shipment Start, but not deliver.");
                                    }
                                }//ProcessMonitor.mPO_Status = PO_Status.DLV
                            }//ProcessMonitor.mPO_Status = PO_Status.PO_Finished;
                        }//ProcessMonitor.mPO_Status = PO_Status.PO_Released;
                    }//ProcessMonitor.mPO_Status = PO_Status.PO_Created;  
                    #endregion
                }//ProcessMonitor.mPO_Status = PO_Status.SO_Created;
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                } 
                #endregion

                #region 统计
                //For ECharts1：PO Create Monitor
                #region For ECharts1：PO Create Monitor
                switch (ProcessMonitor.mPO_Status)
                {
                    case PO_Status.SO_Created:
                        break;
                    case PO_Status.PO_Created:
                    case PO_Status.PO_Released:
                    case PO_Status.PO_Finished:
                    case PO_Status.DLV:
                        //统计SO_CreatedOn为今年的订单                
                        if (ProcessMonitor.mActualTime.Date_SO_CreatedOn.Year == DateTime.Today.Year)
                        {
                            //For ECharts1
                            Data2Trace.count_PO_soCreated_perMonth[ProcessMonitor.mActualTime.Date_SO_CreatedOn.Month - 1]++;
                            Data2Trace.count_PO_soCreated_YTD++;
                            //For ECharts1：1个工作日未创建PO，Avoid system issue：Credit block
                            if (ProcessMonitor.Gap_SO2PO > 1)
                            {
                                Data2Trace.count_CreatedDelay_perMonth[ProcessMonitor.mActualTime.Date_SO_CreatedOn.Month - 1]++;
                                Data2Trace.count_CreatedDelay_YTD++;
                            }
                        }
                        break;
                } 
                #endregion
                //For ECharts2：PO Release Monitor
                #region For ECharts2：PO Release Monitor
                switch (ProcessMonitor.mPO_Status)
                {
                    case PO_Status.SO_Created:
                    case PO_Status.PO_Created:
                        break;
                    case PO_Status.PO_Released:
                    case PO_Status.PO_Finished:
                    case PO_Status.DLV:
                        //统计PO_CreatedOn为今年的订单                
                        if (ProcessMonitor.mActualTime.Date_PO_CreatedOn.Year == DateTime.Today.Year)
                        {
                            //For ECharts2
                            Data2Trace.count_PO_poCreated_perMonth[ProcessMonitor.mActualTime.Date_PO_CreatedOn.Month - 1]++;
                            Data2Trace.count_PO_poCreated_YTD++;
                            //For ECharts2：5个工作日未释放
                            if (ProcessMonitor.Gap_PO2Release > 5)
                            {
                                Data2Trace.count_ReleaseDelay_perMonth[ProcessMonitor.mActualTime.Date_PO_CreatedOn.Month - 1]++;
                                Data2Trace.count_ReleaseDelay_YTD++;
                            }
                        }
                        break;
                } 
                #endregion
                //For ECharts3:DCR(Delivery Class Reliablity)及时交付率
                #region For ECharts3:DCR(Delivery Class Reliablity)及时交付率
                //0400/0481
                switch (ProcessMonitor.mDLV_Plant)
                {
                    case DLV_Plant.P_0400:
                        if (ProcessMonitor.mPO_Status == PO_Status.PO_Finished || ProcessMonitor.mPO_Status == PO_Status.DLV)
                        {
                            Data2Trace.count_0400_PO_DLV++;
                            if (ProcessMonitor.mActualTime.Date_ActualFinishDate > ProcessMonitor.Date_QuotationLT)
                            {
                                //超期完成的订单
                                Data2Trace.count_0400_PO_DLV_Delay++;
                            }
                            else
                            {
                                //及时完成的订单
                                Data2Trace.count_0400_PO_DLV_OnTime++;
                                if (ProcessMonitor.mActualTime.Date_ActualFinishDate.Year == DateTime.Today.Year)
                                {
                                    Data2Trace.count_0400_PO_OnTime_perMonth[ProcessMonitor.mActualTime.Date_ActualFinishDate.Month - 1]++;
                                    Data2Trace.count_0400_PO_OnTime_YTD++;
                                }
                            }
                        }
                        else
                        {
                            Data2Trace.count_0400_PO_noDLV++;
                            if (bool_noLT)
                            {
                                //没有承诺交期，这样状态的订单很安全
                                Data2Trace.count_0400_PO_noDLV_OK++;
                            }
                            else
                            {
                                if (DateTime.Today > ProcessMonitor.Date_QuotationLT)
                                {
                                    //未完成已超期的订单
                                    Data2Trace.count_0400_PO_noDLV_Delay++;
                                    if ((DateTime.Today.AddDays(-1)) == ProcessMonitor.Date_QuotationLT && (ProcessMonitor.mPO_Status != PO_Status.DLV || ProcessMonitor.mPO_Status != PO_Status.PO_Finished))
                                    {
                                        Data2Trace.count_0400_PO_LastFailed++;
                                    }
                                }
                                else
                                {
                                    if (bool_noBasicEndDate)
                                    {
                                        //还没预估计划完成时间，此订单应该不会超期
                                        Data2Trace.count_0400_PO_noDLV_OK++;
                                    }
                                    else
                                    {
                                        if (ProcessMonitor.mEstimatedTime.Date_BasicEndDate > ProcessMonitor.Date_QuotationLT)
                                        {
                                            //未完成预计超期的订单
                                            Data2Trace.count_0400_PO_noDLV_Warning++;
                                        }
                                        else
                                        {
                                            //未完成暂无超期风险的订单
                                            Data2Trace.count_0400_PO_noDLV_OK++;
                                        }
                                    }
                                }
                            }
                        }
                        break;
                    case DLV_Plant.P_0481:
                        if (ProcessMonitor.mPO_Status == PO_Status.PO_Finished || ProcessMonitor.mPO_Status == PO_Status.DLV)
                        {
                            Data2Trace.count_0481_PO_DLV++;
                            if (ProcessMonitor.mActualTime.Date_ActualFinishDate > ProcessMonitor.Date_QuotationLT)
                            {
                                //超期完成的订单
                                Data2Trace.count_0481_PO_DLV_Delay++;
                            }
                            else
                            {
                                //及时完成的订单
                                Data2Trace.count_0481_PO_DLV_OnTime++;
                                if (ProcessMonitor.mActualTime.Date_ActualFinishDate.Year == DateTime.Today.Year)
                                {
                                    Data2Trace.count_0481_PO_OnTime_perMonth[ProcessMonitor.mActualTime.Date_ActualFinishDate.Month - 1]++;
                                    Data2Trace.count_0481_PO_OnTime_YTD++;
                                }
                            }
                        }
                        else
                        {
                            Data2Trace.count_0481_PO_noDLV++;
                            if (bool_noLT)
                            {
                                //没有承诺交期，这样状态的订单很安全
                                Data2Trace.count_0481_PO_noDLV_OK++;
                            }
                            else
                            {
                                if (DateTime.Today > ProcessMonitor.Date_QuotationLT)
                                {
                                    //未完成已超期的订单
                                    Data2Trace.count_0481_PO_noDLV_Delay++;
                                    if ((DateTime.Today.AddDays(-1)) == ProcessMonitor.Date_QuotationLT && (ProcessMonitor.mPO_Status != PO_Status.DLV || ProcessMonitor.mPO_Status != PO_Status.PO_Finished))
                                    {
                                        Data2Trace.count_0481_PO_LastFailed++;
                                    }
                                }
                                else
                                {
                                    if (bool_noBasicEndDate)
                                    {
                                        //还没预估计划完成时间，此订单应该不会超期
                                        Data2Trace.count_0481_PO_noDLV_OK++;
                                    }
                                    else
                                    {
                                        if (ProcessMonitor.mEstimatedTime.Date_BasicEndDate > ProcessMonitor.Date_QuotationLT)
                                        {
                                            //未完成预计超期的订单
                                            Data2Trace.count_0481_PO_noDLV_Warning++;
                                        }
                                        else
                                        {
                                            //未完成暂无超期风险的订单
                                            Data2Trace.count_0481_PO_noDLV_OK++;
                                        }
                                    }
                                }
                            }
                        }
                        break;
                }  
                #endregion
                //For ECharts4:Finish POs
                #region For ECharts4:Finish POs
                //统计ActualFinishDate为今年的月份
                if ((ProcessMonitor.mPO_Status == PO_Status.DLV || ProcessMonitor.mPO_Status == PO_Status.PO_Finished) && ProcessMonitor.mActualTime.Date_ActualFinishDate.Year == DateTime.Today.Year)
                {
                    //For ECharts4
                    Data2Trace.count_PO_Finished_perMonth[ProcessMonitor.mActualTime.Date_ActualFinishDate.Month - 1]++;
                    Data2Trace.count_PO_Finished_YTD++;
                    switch (ProcessMonitor.mDLV_Plant)
                    {
                        case DLV_Plant.P_0400:
                            Data2Trace.count_0400_PO_Finished_perMonth[ProcessMonitor.mActualTime.Date_ActualFinishDate.Month - 1]++;
                            Data2Trace.count_0400_PO_Finished_YTD++;
                            break;
                        case DLV_Plant.P_0481:
                            Data2Trace.count_0481_PO_Finished_perMonth[ProcessMonitor.mActualTime.Date_ActualFinishDate.Month - 1]++;
                            Data2Trace.count_0481_PO_Finished_YTD++;
                            break;
                    }
                    Data2Trace.NetValue_PO_Finished_perMonth[ProcessMonitor.mActualTime.Date_ActualFinishDate.Month - 1] = Data2Trace.NetValue_PO_Finished_perMonth[ProcessMonitor.mActualTime.Date_ActualFinishDate.Month - 1] + ProcessMonitor.mNetValue;
                    Data2Trace.NetValue_PO_Finished_YTD = Data2Trace.NetValue_PO_Finished_YTD + ProcessMonitor.mNetValue;
                }
                else
                {
                    //未完成订单，查看Basic End Date看看预计完成时间在哪个月
                    if (ProcessMonitor.mEstimatedTime.Date_BasicEndDate.Year == DateTime.Today.Year)
                    {
                        //For ECharts4
                        Data2Trace.count_PO_ForecastFinished_perMonth[ProcessMonitor.mEstimatedTime.Date_BasicEndDate.Month - 1]++;
                        Data2Trace.count_PO_ForecastFinished_YTD++;
                    }
                }
                #endregion
                //For ECharts5:SO Create perMonth
                #region For ECharts5:SO Create perMonth
                //统计SO_CreatedOn为今年的月份
                if (ProcessMonitor.mActualTime.Date_SO_CreatedOn.Year == DateTime.Today.Year)
                {
                    //For ECharts5
                    mylist_SO_perMonth[ProcessMonitor.mActualTime.Date_SO_CreatedOn.Month - 1].Add(ProcessMonitor.mSO);
                    mylist_SO_YTD.Add(ProcessMonitor.mSO);
                    Data2Trace.NetValue_SO_soCreated_perMonth[ProcessMonitor.mActualTime.Date_SO_CreatedOn.Month - 1] =
                        Data2Trace.NetValue_SO_soCreated_perMonth[ProcessMonitor.mActualTime.Date_SO_CreatedOn.Month - 1] + 
                        ProcessMonitor.mNetValue;
                    Data2Trace.NetValue_SO_soCreated_YTD = Data2Trace.NetValue_SO_soCreated_YTD + ProcessMonitor.mNetValue;
             
                }
                #endregion
                //For ECharts6:LT(Average Lead Time：CDS)平均交货时间
                #region For ECharts6:LT(Average Lead Time：CDS)平均交货时间
                if (ProcessMonitor.mPO_Status == PO_Status.PO_Finished || ProcessMonitor.mPO_Status == PO_Status.DLV)
                {
                    if (ProcessMonitor.mActualTime.Date_ActualFinishDate.Year == DateTime.Today.Year)
                    {
                        mylist_LT_perMonth[ProcessMonitor.mActualTime.Date_ActualFinishDate.Month - 1].Add(ProcessMonitor.LT);
                        mylist_LT_YTD.Add(ProcessMonitor.LT);
                        //0400/0481
                        switch (ProcessMonitor.mDLV_Plant)
                        {
                            case DLV_Plant.P_0400:
                                mylist_0400_LT_perMonth[ProcessMonitor.mActualTime.Date_ActualFinishDate.Month - 1].Add(ProcessMonitor.LT);
                                mylist_0400_LT_YTD.Add(ProcessMonitor.LT);
                                break;
                            case DLV_Plant.P_0481:
                                mylist_0481_LT_perMonth[ProcessMonitor.mActualTime.Date_ActualFinishDate.Month - 1].Add(ProcessMonitor.LT);
                                mylist_0481_LT_YTD.Add(ProcessMonitor.LT);
                                break;
                        }
                    }
                }
                
                //public static double[] LT_perMonth = new double[12];
                //public static double[] LT_perWeek = new double[4];
                //public static double LT_YTD;
                //public static double[] LT_0400_perMonth = new double[12];
                //public static double[] LT_0400_perWeek = new double[4];
                //public static double LT_0400_YTD;
                //public static double[] LT_0481_perMonth = new double[12];
                //public static double[] LT_0481_perWeek = new double[4];
                //public static double LT_0481_YTD; 
	            #endregion
                //For ECharts7:Reminders for 3Weeks&2Weeks
                #region For ECharts7:Reminders for 3Weeks&2Weeks
                //public static int count_Reminder3Weeks_LT;
                //public static int count_Reminder2Weeks_LT;
                //public static int count_Reminder3Weeks_ConfirmedDt;
                //public static int count_Reminder2Weeks_ConfirmedDt;
                //0400未完成的订单
                if ((ProcessMonitor.mPO_Status != PO_Status.DLV && ProcessMonitor.mPO_Status != PO_Status.PO_Finished) && ProcessMonitor.mDLV_Plant == DLV_Plant.P_0400)
                {
                    if (ProcessMonitor.Date_QuotationLT >DateTime.Today.AddDays(7)
                        &&ProcessMonitor.Date_QuotationLT <=DateTime.Today.AddDays(14))
                    {
                        Data2Trace.count_Reminder2Weeks_LT++;
                    }
                    else if (ProcessMonitor.Date_QuotationLT >DateTime.Today.AddDays(14)
                        &&ProcessMonitor.Date_QuotationLT <=DateTime.Today.AddDays(21))
                    {
                        Data2Trace.count_Reminder3Weeks_LT++;
                    }
                    if (ProcessMonitor.mEstimatedTime.Date_1stConfirmedDt > DateTime.Today.AddDays(7) 
                        && ProcessMonitor.mEstimatedTime.Date_1stConfirmedDt <= DateTime.Today.AddDays(14))
                    {
                        Data2Trace.count_Reminder2Weeks_ConfirmedDt++;
                    }
                    else if (ProcessMonitor.mEstimatedTime.Date_1stConfirmedDt > DateTime.Today.AddDays(14) 
                        && ProcessMonitor.mEstimatedTime.Date_1stConfirmedDt <= DateTime.Today.AddDays(21))
                    {
                        Data2Trace.count_Reminder3Weeks_ConfirmedDt++;
                    }
                }
                #endregion
                //For ECharts8:Repertory Monitor
                #region For ECharts8:Repertory Monitor
                if (ProcessMonitor.mPO_Status == PO_Status.PO_Finished)
                {
                    Data2Trace.count_Stored++;
                }
                else if (ProcessMonitor.mPO_Status == PO_Status.DLV)
                {
                    if (ProcessMonitor.Gap_Finish2Shipment <= 5)
                    {
                        Data2Trace.count_Stored_5++;
                    }
                    else if (ProcessMonitor.Gap_Finish2Shipment <= 10)
                    {
                        Data2Trace.count_Stored_5_10++;
                    }
                    else if (ProcessMonitor.Gap_Finish2Shipment <= 20)
                    {
                        Data2Trace.count_Stored_10_20++;
                    }
                    else
                    {
                        Data2Trace.count_Stored_20++;
                    }
                }
                #endregion              
                #endregion
                #endregion
            }
            return true;
        }
        catch (Exception)
        {
            return false;
        }        
    }
}