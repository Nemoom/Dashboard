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
        excelHead.Add(new KeyValuePair<int, string>( 0, "SOrg."));        
        excelHead.Add(new KeyValuePair<int, string>( 1, "SaTy"));        
        excelHead.Add(new KeyValuePair<int, string>( 2, "Sales Doc."));            
        excelHead.Add(new KeyValuePair<int, string>( 3, "SalesItem"));            
        excelHead.Add(new KeyValuePair<int, string>( 4, "DLV Plant"));            
        excelHead.Add(new KeyValuePair<int, string>( 5, "Prod Ord Type"));                
        excelHead.Add(new KeyValuePair<int, string>( 6, "Production Order"));                    
        excelHead.Add(new KeyValuePair<int, string>( 7, "Status"));        
        excelHead.Add(new KeyValuePair<int, string>( 8, "Material"));            
        excelHead.Add(new KeyValuePair<int, string>( 9, "Sold-To Pt"));            
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

        excelHeader[ 0]="SOrg.";
        excelHeader[ 1]="SaTy";
        excelHeader[ 2]="Sales Doc.";
        excelHeader[ 3]="SalesItem";
        excelHeader[ 4]="DLV Plant";
        excelHeader[ 5]="Prod Ord Type";
        excelHeader[ 6]="Production Order";
        excelHeader[ 7]="Status";
        excelHeader[ 8]="Material";
        excelHeader[ 9]="Sold-To Pt";
        excelHeader[10]= "Sold-To Party";
        excelHeader[11]= "Segment";
        excelHeader[12]= "Personnel";
        excelHeader[13]= "Sales Responsible";
        excelHeader[14]= "SOff.";
        excelHeader[15]= "Sales Office";
        excelHeader[16]= "Product Hierarchy";
        excelHeader[17]= "WBS Element";
        excelHeader[18]= "WBS Element";
        excelHeader[19]= "Project Description";
        excelHeader[20]= "Proj Engineer";
        excelHeader[21]= "Proj Engineer Name";
        excelHeader[22]= "Net value";
        excelHeader[23]= "Order Quantity";
        excelHeader[24]= "Confirmed Qty";
        excelHeader[25]= "SO CreatedOn";
        excelHeader[26]= "Prod CreatedOn";
        excelHeader[27]= "Actual Release Date";
        excelHeader[28]= "Actual Finish Date";
        excelHeader[29]= "1st GR for Prod Order";
        excelHeader[30]= "Quotation LT";
        excelHeader[31]= "1st ConfirmedDt";
        excelHeader[32]= "Requested Date";
        excelHeader[33]= "Basic End Date";
        excelHeader[34]= "ShipmentStartOn";
        excelHeader[35]= "Route";
        excelHeader[36]= "Route";
        excelHeader[37]= "TransitTime";
        excelHeader[38]= "Delivery Number";
        excelHeader[39]= "Shipment Number";
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
        sht = (XSSFSheet)wb.GetSheet(sheetName);
        //取行Excel的最大行数
        rowsCount = sht.PhysicalNumberOfRows;
        //取所有行中的最大列数
        colsCount = sht.GetRow(0).PhysicalNumberOfCells;
        for (int i = 0; i < colsCount; i++)
        {
            if (((XSSFCell)((XSSFRow)sht.GetRow(0)).GetCell(i)).StringCellValue != excelHeader[i])
            { 
                bool_ImportResult = false;
                break;
            }
        }
        return bool_ImportResult;
    }

    public void RowRecordProcess(int rowIndex)
    {
        XSSFRow nRow = (XSSFRow)sht.GetRow(rowIndex);
        ProcessMonitor.mActualTime.Date_SO_CreatedOn = DateTime.FromOADate(((XSSFCell)nRow.GetCell(25)).NumericCellValue);
        ProcessMonitor.mActualTime.Date_PO_CreatedOn = DateTime.FromOADate(((XSSFCell)nRow.GetCell(26)).NumericCellValue);
        ProcessMonitor.mActualTime.Date_ActualReleaseDate = DateTime.FromOADate(((XSSFCell)nRow.GetCell(27)).NumericCellValue);
        ProcessMonitor.mActualTime.Date_ActualFinishDate = DateTime.FromOADate(((XSSFCell)nRow.GetCell(28)).NumericCellValue);
        ProcessMonitor.mActualTime.Date_1stGRforPO = DateTime.FromOADate(((XSSFCell)nRow.GetCell(29)).NumericCellValue);
        ProcessMonitor.mActualTime.Date_ShipmentStartOn = DateTime.FromOADate(((XSSFCell)nRow.GetCell(39)).NumericCellValue);

        ProcessMonitor.mEstimatedTime.QuotationLT = (int)((XSSFCell)nRow.GetCell(30)).NumericCellValue;
        ProcessMonitor.mEstimatedTime.Date_1stConfirmedDt = DateTime.FromOADate(((XSSFCell)nRow.GetCell(31)).NumericCellValue);
        ProcessMonitor.mEstimatedTime.Date_RequestDate = DateTime.FromOADate(((XSSFCell)nRow.GetCell(32)).NumericCellValue);
        ProcessMonitor.mEstimatedTime.Date_BasicEndDate = DateTime.FromOADate(((XSSFCell)nRow.GetCell(33)).NumericCellValue);
    }
}