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
        //wb = new XSSFWorkbook(File.OpenRead(@"C:\Users\Public\Music\" + filename));
        wb = new XSSFWorkbook(File.OpenRead(@"F:\" + filename));
        sht = (XSSFSheet)wb.GetSheet(sheetName);
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
        for (int i = 1; i < rowsCount; i++)
        {
            RowRecordProcess(i);
        }
    }
    List<string>[] mylist_SO_perMonth = new List<string>[12];
    List<string> mylist_SO_YTD = new List<string>();
    List<int>[] mylist_LT_perMonth = new List<int>[12];
    List<int>[] mylist_LT_perWeek = new List<int>[4];
    List<int> mylist_LT_YTD = new List<int>();
    public void RowRecordProcess(int rowIndex)
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
            ProcessMonitor.mSO = ((XSSFCell)nRow.GetCell(2)).StringCellValue;//每单都有
            ProcessMonitor.mPO = ((XSSFCell)nRow.GetCell(6)).StringCellValue;//未创建生产单号的此项为空
            if (((XSSFCell)nRow.GetCell(7)).StringCellValue == "DIV")
            {
                #region 已交付
                //已交付，应该数据齐活儿
                ProcessMonitor.mPO_Status = PO_Status.DIV;
                Data2Trace.count_PO_DIV++;
                try
                {
                    ProcessMonitor.mActualTime.Date_SO_CreatedOn = DateTime.FromOADate(((XSSFCell)nRow.GetCell(25)).NumericCellValue);
                    ProcessMonitor.mNetValue = ((XSSFCell)nRow.GetCell(22)).NumericCellValue;
                    //统计今年的月份
                    if (ProcessMonitor.mActualTime.Date_SO_CreatedOn.Year == DateTime.Today.Year)
                    {
                        //For ECharts5
                        mylist_SO_perMonth[ProcessMonitor.mActualTime.Date_SO_CreatedOn.Month - 1].Add(ProcessMonitor.mSO);
                        mylist_SO_YTD.Add(ProcessMonitor.mSO);
                    }
                    ProcessMonitor.mActualTime.Date_PO_CreatedOn = DateTime.FromOADate(((XSSFCell)nRow.GetCell(26)).NumericCellValue);
                    //统计今年的月份                
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
                    ProcessMonitor.mActualTime.Date_ActualReleaseDate = DateTime.FromOADate(((XSSFCell)nRow.GetCell(27)).NumericCellValue);
                    if (ProcessMonitor.mActualTime.Date_PO_CreatedOn.Year == DateTime.Today.Year)
                    {
                        //For ECharts2
                        Data2Trace.count_PO_poCreated_perMonth[ProcessMonitor.mActualTime.Date_PO_CreatedOn.Month - 1]++;
                        Data2Trace.count_PO_poCreated_YTD++;
                        //For ECharts2：5个工作日未释放
                        if (ProcessMonitor.Gap_PO2Release > 5)
                        {
                            Data2Trace.count_ReleaseDelay_perMonth[ProcessMonitor.mActualTime.Date_SO_CreatedOn.Month - 1]++;
                            Data2Trace.count_ReleaseDelay_YTD++;
                        }
                    }
                    ProcessMonitor.mActualTime.Date_ActualFinishDate = DateTime.FromOADate(((XSSFCell)nRow.GetCell(28)).NumericCellValue);
                    //统计今年的月份
                    if (ProcessMonitor.mActualTime.Date_ActualFinishDate.Year == DateTime.Today.Year)
                    {
                        //For ECharts4
                        Data2Trace.count_PO_Finished_perMonth[ProcessMonitor.mActualTime.Date_ActualFinishDate.Month - 1]++;
                        Data2Trace.count_PO_Finished_YTD++;
                        Data2Trace.NetValue_PO_Finished_perMonth[ProcessMonitor.mActualTime.Date_ActualFinishDate.Month - 1] = Data2Trace.NetValue_PO_Finished_perMonth[ProcessMonitor.mActualTime.Date_ActualFinishDate.Month - 1] + ProcessMonitor.mNetValue;
                        Data2Trace.NetValue_PO_Finished_YTD = Data2Trace.NetValue_PO_Finished_YTD + ProcessMonitor.mNetValue;
                    }
                    ProcessMonitor.mActualTime.Date_1stGRforPO = DateTime.FromOADate(((XSSFCell)nRow.GetCell(29)).NumericCellValue);
                    ProcessMonitor.mActualTime.Date_ShipmentStartOn = DateTime.FromOADate(((XSSFCell)nRow.GetCell(34)).NumericCellValue);
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

                    ProcessMonitor.mEstimatedTime.QuotationLT = (int)((XSSFCell)nRow.GetCell(30)).NumericCellValue;
                    ProcessMonitor.mEstimatedTime.Date_1stConfirmedDt = DateTime.FromOADate(((XSSFCell)nRow.GetCell(31)).NumericCellValue);
                    ProcessMonitor.mEstimatedTime.Date_RequestDate = DateTime.FromOADate(((XSSFCell)nRow.GetCell(32)).NumericCellValue);
                    //已完成订单，Date_BasicEndDate忽略
                    ProcessMonitor.mEstimatedTime.Date_BasicEndDate = DateTime.FromOADate(((XSSFCell)nRow.GetCell(33)).NumericCellValue);

                    if (ProcessMonitor.mActualTime.Date_ActualFinishDate > ProcessMonitor.Date_QuotationLT)
                    {
                        //超期交付的订单
                        Data2Trace.count_PO_DIV_Delay++;
                    }
                    else
                    {
                        //及时交付的订单
                        Data2Trace.count_PO_DIV_OnTime++;
                    }

                }
                catch (Exception)
                {
                    Console.Write("Incomplete data");
                }
                #endregion
            }
            else
            {
                #region 未交付
                //未交付，部分时间节点应该还为空
                Data2Trace.count_PO_noDIV++;
                try
                {
                    bool bool_noLT, bool_noBasicEndDate;
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
                            #region Released PO
                            //统计今年的月份                
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

                            ProcessMonitor.mActualTime.Date_ActualReleaseDate = DateTime.FromOADate(((XSSFCell)nRow.GetCell(27)).NumericCellValue);
                            if (ProcessMonitor.mActualTime.Date_PO_CreatedOn.Year == DateTime.Today.Year)
                            {
                                //For ECharts2
                                Data2Trace.count_PO_poCreated_perMonth[ProcessMonitor.mActualTime.Date_PO_CreatedOn.Month - 1]++;
                                Data2Trace.count_PO_poCreated_YTD++;
                                //For ECharts2：5个工作日未释放
                                if (ProcessMonitor.Gap_PO2Release > 5)
                                {
                                    Data2Trace.count_ReleaseDelay_perMonth[ProcessMonitor.mActualTime.Date_SO_CreatedOn.Month - 1]++;
                                    Data2Trace.count_ReleaseDelay_YTD++;
                                }
                            }

                            if (((XSSFCell)nRow.GetCell(28)).CellType == NPOI.SS.UserModel.CellType.Blank)
                            {
                                //PO not Finished
                                ProcessMonitor.mPO_Status = PO_Status.PO_Released;
                                //未完成订单，查看Basic End Date看看预计完成时间在哪个月
                                if (ProcessMonitor.mEstimatedTime.Date_BasicEndDate.Year == DateTime.Today.Year)
                                {
                                    //For ECharts4
                                    Data2Trace.count_PO_ForecastFinished_perMonth[ProcessMonitor.mEstimatedTime.Date_BasicEndDate.Month - 1]++;
                                    Data2Trace.count_PO_ForecastFinished_YTD++;
                                }
                            }
                            else
                            {
                                #region 已完成订单
                                ProcessMonitor.mActualTime.Date_ActualFinishDate = DateTime.FromOADate(((XSSFCell)nRow.GetCell(28)).NumericCellValue);
                                ProcessMonitor.mActualTime.Date_1stGRforPO = DateTime.FromOADate(((XSSFCell)nRow.GetCell(29)).NumericCellValue);
                                //统计今年的月份
                                if (ProcessMonitor.mActualTime.Date_ActualFinishDate.Year == DateTime.Today.Year)
                                {
                                    //For ECharts4，已完成但未快递给客户的订单
                                    Data2Trace.count_PO_Finished_perMonth[ProcessMonitor.mActualTime.Date_ActualFinishDate.Month - 1]++;
                                    Data2Trace.count_PO_Finished_YTD++;
                                    Data2Trace.NetValue_PO_Finished_perMonth[ProcessMonitor.mActualTime.Date_ActualFinishDate.Month - 1] = Data2Trace.NetValue_PO_Finished_perMonth[ProcessMonitor.mActualTime.Date_ActualFinishDate.Month - 1] + ProcessMonitor.mNetValue;
                                    Data2Trace.NetValue_PO_Finished_YTD = Data2Trace.NetValue_PO_Finished_YTD + ProcessMonitor.mNetValue;
                                }
                                if (((XSSFCell)nRow.GetCell(34)).CellType == NPOI.SS.UserModel.CellType.Blank)
                                {
                                    //PO not Shipment
                                    ProcessMonitor.mPO_Status = PO_Status.PO_Finished;
                                    Data2Trace.count_Stored++;
                                }
                                else
                                {
                                    ProcessMonitor.mActualTime.Date_ShipmentStartOn = DateTime.FromOADate(((XSSFCell)nRow.GetCell(34)).NumericCellValue);
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
                            }
                            #endregion
                        }

                    }

                    if (bool_noLT)
                    {
                        //没有承诺交期，这样状态的订单很安全
                        Data2Trace.count_PO_noDIV_OK++;
                    }
                    else
                    {
                        if (DateTime.Today > ProcessMonitor.Date_QuotationLT)
                        {
                            //未完成已超期的订单
                            Data2Trace.count_PO_noDIV_Delay++;
                            if ((DateTime.Today.AddDays(-1)) == ProcessMonitor.Date_QuotationLT && (ProcessMonitor.mPO_Status != PO_Status.DIV || ProcessMonitor.mPO_Status != PO_Status.PO_Finished))
                            {
                                Data2Trace.count_PO_LastFailed++;
                            }
                        }
                        else
                        {
                            if (bool_noBasicEndDate)
                            {
                                //还没预估计划完成时间，此订单应该不会超期
                                Data2Trace.count_PO_noDIV_OK++;
                            }
                            else
                            {
                                if (ProcessMonitor.mEstimatedTime.Date_BasicEndDate > ProcessMonitor.Date_QuotationLT)
                                {
                                    //未完成预计超期的订单
                                    Data2Trace.count_PO_noDIV_Warning++;
                                }
                                else
                                {
                                    //未完成暂无超期风险的订单
                                    Data2Trace.count_PO_noDIV_OK++;
                                }
                            }
                        }
                    }

                }
                catch (Exception)
                {

                    throw;
                }
                #endregion
            }
            //统计数据   
            #endregion         

        }
        
    }
}