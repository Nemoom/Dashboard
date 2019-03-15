<%@ WebHandler Language="C#" Class="Handler" %>

using System;
using System.Web;
using System.IO;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using NPOI.SS.Util;
using NPOI.Util;
using NPOI.XSSF.UserModel;

public class Handler : IHttpHandler {
    
    public void ProcessRequest (HttpContext context) {
        context.Response.ContentType = "text/plain";
        //context.Response.Write("Hello World");
        string filename = context.Request.Form["filename"];
        if (filename!=null)
        {
            GetInfoFromXLSX(filename);
            string json = "";
            json += "[";
            json += "{\"name\": \"已交付\",\"value\": " + (num_ok + num_nok) + "},";
            json += "{\"name\": \"未交付\",\"value\": " + num_noDIV + "},";
            json += "{\"name\": \"延期交付\",\"value\": " + num_nok + "},";
            json += "{\"name\": \"正常交付\",\"value\": " + num_ok + "},";
            json += "{\"name\": \"正常\",\"value\": " + num_ok_noDIV + "},";
            json += "{\"name\": \"已超期\",\"value\": " + num_Error_noDIV + "},";
            json += "{\"name\": \"即将超期\",\"value\": " + num_Warning_noDIV + "}";
            json += "]";
            context.Response.Write(json);
        }
        
    }

    /// <summary>
    /// 把数据从Excel装载到DataTable
    /// </summary>
    /// <param name="pathName">带路径的Excel文件名</param>
    /// <param name="sheetName">工作表名</param>
    /// <param name="tbContainer">将数据存入的DataTable</param>
    /// <returns></returns>
    public DataTable ExcelToDataTable(string pathName, string sheetName)
    {
        DataTable tbContainer = new DataTable();
        string strConn = string.Empty;
        if (string.IsNullOrEmpty(sheetName)) { sheetName = "Sheet1"; }
        FileInfo file = new FileInfo(pathName);
        if (!file.Exists) { throw new Exception("文件不存在"); }
        string extension = file.Extension;
        switch (extension)
        {
            case ".xls":
                strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathName + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'";
                break;
            case ".xlsx":
                strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathName + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1;'";
                break;
            default:
                strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathName + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'";
                break;
        }
        //链接Excel
        OleDbConnection cnnxls = new OleDbConnection(strConn);
        //读取Excel里面有 表Sheet1
        OleDbDataAdapter oda = new OleDbDataAdapter(string.Format("select * from [{0}$]", sheetName), cnnxls);
        DataSet ds = new DataSet();
        //将Excel里面有表内容装载到内存表中！
        oda.Fill(tbContainer);
        return tbContainer;
    }
 
    public bool IsReusable {
        get {
            return false;
        }
    }

    private XSSFSheet sht;
    protected String excelContent;
    int num_noDIV = 0;
    int num_ok = 0;
    int num_nok = 0;
    int num_ok_noDIV = 0;
    int num_Error_noDIV = 0;
    int num_Warning_noDIV = 0;

    private void GetInfoFromXLSX(string filename)
    {
        XSSFWorkbook wb = new XSSFWorkbook(File.OpenRead(@"C:\Users\Public\Music\"+filename));
        sht = (XSSFSheet)wb.GetSheet("Sheet1");
        //取行Excel的最大行数
        int rowsCount = sht.PhysicalNumberOfRows;
        //为保证Table布局与Excel一样，这里应该取所有行中的最大列数（需要遍历整个Sheet）。
        //为少一交全Excel遍历，提高性能，我们可以人为把第0行的列数调整至所有行中的最大列数。
        int colsCount = sht.GetRow(0).PhysicalNumberOfCells;

        int colSpan;
        int rowSpan;
        bool isByRowMerged;

        StringBuilder table = new StringBuilder(rowsCount * 32);

        table.Append("<table border='1px'>");
        for (int rowIndex = 0; rowIndex < rowsCount; rowIndex++)
        {
            XSSFRow nRow = (XSSFRow)sht.GetRow(rowIndex);
            XSSFCell nCell_Status = (XSSFCell)nRow.GetCell(7);
            XSSFCell nCell_Create = (XSSFCell)nRow.GetCell(25);
            XSSFCell nCell_Finish = (XSSFCell)nRow.GetCell(28);
            XSSFCell nCell_LT = (XSSFCell)nRow.GetCell(30);
            XSSFCell nCell_Confirmed = (XSSFCell)nRow.GetCell(31);
            DateTimeFormatInfo dtFormat = new DateTimeFormatInfo();
            dtFormat.ShortDatePattern = "yyyy/MM/dd";
            if (nCell_Create.CellType == NPOI.SS.UserModel.CellType.String)
            {
                continue;//跳过表头那一行
            }
            switch (nCell_Status.CellType)
            {
                case NPOI.SS.UserModel.CellType.Unknown:
                    break;
                case NPOI.SS.UserModel.CellType.Numeric:
                    break;
                case NPOI.SS.UserModel.CellType.String:
                    break;
                case NPOI.SS.UserModel.CellType.Formula:
                    break;
                case NPOI.SS.UserModel.CellType.Blank:
                    break;
                case NPOI.SS.UserModel.CellType.Boolean:
                    break;
                case NPOI.SS.UserModel.CellType.Error:
                    break;
                default:
                    break;
            }
            if (nCell_Status.StringCellValue != "DLV")
            {
                num_noDIV++;//未交付
                try
                {
                    DateTime dt;
                    dt = Convert.ToDateTime(nCell_Confirmed.DateCellValue);
                    if ((DateTime.Today - dt).Days > 7)
                    {
                        num_ok_noDIV++;//正常
                    }
                    else if ((DateTime.Today - dt).Days < 0)
                    {
                        num_Error_noDIV++;//已超期
                    }
                    else
                    {
                        num_Warning_noDIV++;//即将超期
                    }
                }
                catch (Exception)
                {
                }

                table.Append("<tr>");

                //for (int colIndex = 0; colIndex < colsCount; colIndex++)
                //{
                //    GetTdMergedInfo(rowIndex, colIndex, out colSpan, out rowSpan, out isByRowMerged);
                //    //如果已经被行合并包含进去了就不输出TD了。
                //    //注意被合并的行或列不输出的处理方式不一样，见下面一处的注释说明了列合并后不输出TD的处理方式。
                //    if (isByRowMerged)
                //    {
                //        continue;
                //    }

                //    table.Append("<td");
                //    if (colSpan > 1)
                //        table.Append(string.Format(" colSpan={0}", colSpan));
                //    if (rowSpan > 1)
                //        table.Append(string.Format(" rowSpan={0}", rowSpan));
                //    table.Append(">");

                //    table.Append(sht.GetRow(rowIndex).GetCell(colIndex));

                //    //列被合并之后此行将少输出colSpan-1个TD。
                //    if (colSpan > 1)
                //        colIndex += colSpan - 1;

                //    table.Append("</td>");

                //}
            }
            else
            {
                DateTime dt_Act;
                DateTime dt_Plan;
                try
                {
                    //dt_Act = Convert.ToDateTime(nCell_Finish.DateCellValue);
                    dt_Act = DateTime.FromOADate(nCell_Finish.NumericCellValue);
                    //dt_Plan = Convert.ToDateTime(nCell_Confirmed.DateCellValue);
                    dt_Plan = DateTime.FromOADate(nCell_Confirmed.NumericCellValue);
                    if (dt_Act > dt_Plan)
                    {
                        num_nok++;
                    }
                    else
                    {
                        num_ok++;
                    }
                }
                catch (Exception)
                {
                }
            }
            table.Append("</tr>");
        }
        table.Append("</table>");
        String mResult = num_ok.ToString() + "," + num_nok.ToString() + "," + num_noDIV.ToString() + "," + num_ok_noDIV.ToString() + "," + num_Warning_noDIV.ToString() + "," + num_Error_noDIV.ToString();

        this.excelContent = table.ToString();
    }

    /// <summary>
    ///  获取Table某个TD合并的列数和行数等信息。与Excel中对应Cell的合并行数和列数一致。
    /// </summary>
    /// <param name="rowIndex">行号</param>
    /// <param name="colIndex">列号</param>
    /// <param name="colspan">TD中需要合并的行数</param>
    /// <param name="rowspan">TD中需要合并的列数</param>
    /// <param name="rowspan">此单元格是否被某个行合并包含在内。如果被包含在内，将不输出TD。</param>
    /// <returns></returns>
    private void GetTdMergedInfo(int rowIndex, int colIndex, out int colspan, out int rowspan, out bool isByRowMerged)
    {
        colspan = 1;
        rowspan = 1;
        isByRowMerged = false;
        int regionsCuont = sht.NumMergedRegions;
        CellRangeAddress region;
        for (int i = 0; i < regionsCuont; i++)
        {
            region = sht.GetMergedRegion(i);
            if (region.FirstRow == rowIndex && region.FirstColumn == colIndex)
            {
                colspan = region.LastColumn - region.FirstColumn + 1;
                rowspan = region.LastRow - region.FirstRow + 1;
                return;
            }
            else if (rowIndex > region.FirstRow && rowIndex <= region.LastRow && colIndex >= region.FirstColumn && colIndex <= region.LastColumn)
            {
                isByRowMerged = true;
            }
        }
    }
}