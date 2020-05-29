using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Globalization;
using NPOI.SS.Util;
using NPOI.Util;
using NPOI.XSSF.UserModel;
using System.Text;
using System.IO;
using System.Diagnostics;

public partial class DashBoard : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        //////this.Page.ClientScript.RegisterStartupScript(this.Page.GetType(), "setLoading", "setLoading()", true);
        //////this.Page.ClientScript.RegisterStartupScript(this.Page.GetType(), "setECharts1", "setECharts1('"+this.Application["filePath"].ToString()+"')", true);
        ////GetInfoFromXLSX(this.Application["filePath"].ToString());
        ////this.Page.ClientScript.RegisterStartupScript(this.Page.GetType(), "setECharts2", "setECharts2('" + this.Application["filePath"].ToString() + "','" + excelResult_POnum + "','" + excelResult_Finishnum + "','" + excelResult_NetValue + "')", true);
        //////this.Page.ClientScript.RegisterStartupScript(this.Page.GetType(), "setECharts2", "setECharts2()", true);
        ExcelOperation mExcelOperation = new ExcelOperation();
        if (File.Exists(@"C:\Users\Public\Music\" + this.Application["filePath_CbyC"].ToString()))
        {
            try
            {
                mExcelOperation.GetCbyCList(this.Application["filePath_CbyC"].ToString());
            }
            catch (Exception)
            {
                
                throw;
            }
        }
        if (File.Exists(@"C:\Users\Public\Music\" + this.Application["filePath_MatchReq"].ToString()))
        {
            try
            {
                mExcelOperation.GetMatchReqList(this.Application["filePath_MatchReq"].ToString());
            }
            catch (Exception)
            {

                throw;
            }
        }
        bool bool_ExcelImpoert = mExcelOperation.ExcelImportWithLayoutCheck(this.Application["filePath"].ToString(), "Sheet1");
        if (bool_ExcelImpoert)
        {           
            mExcelOperation.TraceFromExcel();
            if (this.Application["filePath"].ToString().Substring(this.Application["filePath"].ToString().Length-8,3)=="SKA"||this.Application["filePath"].ToString().Substring(this.Application["filePath"].ToString().Length-8,3)=="FKA")
            {
                lbl_DateInfo.Text = this.Application["filePath"].ToString().Substring(this.Application["filePath"].ToString().Length-8,3)+this.Application["filePath"].ToString().Substring(0,8);

            }
            else
            {
                lbl_DateInfo.Text = this.Application["filePath"].ToString().Substring(0, 8);

            }
            this.Page.ClientScript.RegisterStartupScript(this.Page.GetType(), "setECharts2", "setECharts2('" +
            Data2Trace.values_ECharts1_1 + "','"
            + Data2Trace.values_ECharts1_2 + "','"
            + Data2Trace.values_ECharts1_3 + "','"
            + Data2Trace.values_ECharts1_4 + "','"
            + Data2Trace.values_ECharts2_1 + "','"
            + Data2Trace.values_ECharts2_2 + "','"
            + Data2Trace.values_ECharts2_3 + "','"
            + Data2Trace.values_ECharts2_4 + "','"
            + Data2Trace.values_ECharts3_1 + "','"
            + Data2Trace.values_ECharts3_2 + "','"
            + Data2Trace.values_ECharts3_3 + "','"
            + Data2Trace.values_ECharts3_4 + "')", true);
            Label1.Text = this.Application["filePath"].ToString();
           
        }
        else
        {
            Response.Write("Failed to read table data. Server may be unavailable or table may have a wrong data format.");
        }       
       
    }

    private XSSFSheet sht;
    protected String excelContent;
    protected String excelResult_POnum;
    protected String excelResult_Finishnum;
    protected String excelResult_NetValue;
    int num_noDIV = 0;
    int num_ok = 0;
    int num_nok = 0;
    int num_ok_noDIV = 0;
    int num_Error_noDIV = 0;
    int num_Warning_noDIV = 0;
    int[] POnum_PerMonth = new int[12];
    int[] Finishnum_PerMonth = new int[12];
    double[] NetValue_PerMonth = new double[12];

    private void GetInfoFromXLSX(string filename)
    {
        try
        {
            ExcelOperation mExcelOperation=new ExcelOperation();

            bool bool_excelload = mExcelOperation.ExcelImportWithLayoutCheck(filename, "Sheet1");
            XSSFWorkbook wb = new XSSFWorkbook(File.OpenRead(@"C:\Users\Public\Music\" + filename));
            sht = (XSSFSheet)wb.GetSheet("Sheet1");
            //取行Excel的最大行数
            int rowsCount = sht.PhysicalNumberOfRows;
            //为保证Table布局与Excel一样，这里应该取所有行中的最大列数（需要遍历整个Sheet）。
            //为少一交全Excel遍历，提高性能，我们可以人为把第0行的列数调整至所有行中的最大列数。
            int colsCount = sht.GetRow(0).PhysicalNumberOfCells;

            StringBuilder table = new StringBuilder(rowsCount * 32);

            table.Append("<table border='1px'>");
            for (int rowIndex = 0; rowIndex < rowsCount; rowIndex++)
            {
                XSSFRow nRow = (XSSFRow)sht.GetRow(rowIndex);
                XSSFCell nCell_Status = (XSSFCell)nRow.GetCell(7);
                XSSFCell nCell_NetValue = (XSSFCell)nRow.GetCell(22);
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
                if (nCell_Finish.CellType == NPOI.SS.UserModel.CellType.Numeric)
                {
                    DateTime dt_Finish;
                    dt_Finish = DateTime.FromOADate(nCell_Finish.NumericCellValue);
                    //try
                    //{
                    //    dt_Finish = Convert.ToDateTime(nCell_Finish.DateCellValue);
                    //}
                    //catch (Exception)
                    //{
                    //    dt_Finish = DateTime.FromOADate(nCell_Finish.NumericCellValue);                    
                    //}

                    if (dt_Finish.Year == 2019 )
                    {
                        Finishnum_PerMonth[dt_Finish.Month - 1]++;
                    }
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
                DateTime dt_SOCreated;
                dt_SOCreated = DateTime.FromOADate(nCell_Create.NumericCellValue);
                //try
                //{
                //    dt_SOCreated = Convert.ToDateTime(nCell_Create.DateCellValue);
                //}
                //catch (Exception)
                //{
                //    dt_SOCreated = DateTime.FromOADate(nCell_Create.NumericCellValue);
                //}
                if (dt_SOCreated.Year == 2019 )
                {
                    POnum_PerMonth[dt_SOCreated.Month - 1]++;
                    NetValue_PerMonth[dt_SOCreated.Month - 1] = NetValue_PerMonth[dt_SOCreated.Month - 1] + Convert.ToDouble(nCell_NetValue.NumericCellValue);
                }

                if (nCell_Status.StringCellValue != "DLV")
                {
                    num_noDIV++;//未交付
                    try
                    {
                        DateTime dt;
                        dt = DateTime.FromOADate(nCell_Confirmed.NumericCellValue);
                        //try
                        //{
                        //    dt = Convert.ToDateTime(nCell_Confirmed.DateCellValue);
                        //}
                        //catch (Exception)
                        //{
                        //    dt = DateTime.FromOADate(nCell_Confirmed.NumericCellValue);
                        //}

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

                }
                else
                {
                    DateTime dt_Act;
                    DateTime dt_Plan;
                    try
                    {
                        dt_Act = DateTime.FromOADate(nCell_Finish.NumericCellValue);
                        //try
                        //{
                        //    dt_Act = Convert.ToDateTime(nCell_Finish.DateCellValue);
                        //}
                        //catch (Exception)
                        //{
                        //    dt_Act = DateTime.FromOADate(nCell_Finish.NumericCellValue);
                        //}
                        dt_Plan = DateTime.FromOADate(nCell_Confirmed.NumericCellValue);
                        //try
                        //{
                        //    dt_Plan = Convert.ToDateTime(nCell_Confirmed.DateCellValue);
                        //}
                        //catch (Exception)
                        //{
                        //    dt_Plan = DateTime.FromOADate(nCell_Confirmed.NumericCellValue);
                        //}
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
            excelResult_POnum = "";
            excelResult_Finishnum = "";
            excelResult_NetValue = "";
            for (int i = 0; i < POnum_PerMonth.Length; i++)
            {
                if (i < POnum_PerMonth.Length - 1)
                {
                    excelResult_POnum = excelResult_POnum + POnum_PerMonth[i].ToString() + ",";
                    excelResult_Finishnum = excelResult_Finishnum + Finishnum_PerMonth[i].ToString() + ",";
                    excelResult_NetValue = excelResult_NetValue + NetValue_PerMonth[i].ToString("F2") + ",";
                }
                else
                {
                    excelResult_POnum = excelResult_POnum + POnum_PerMonth[i].ToString();
                    excelResult_Finishnum = excelResult_Finishnum + Finishnum_PerMonth[i].ToString();
                    excelResult_NetValue = excelResult_NetValue + NetValue_PerMonth[i].ToString("F2");
                }
            }
            this.excelContent = table.ToString();
        }
        catch (Exception)
        {
            
            throw;
        }
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
    protected void DropDownList1_SelectedIndexChanged(object sender, EventArgs e)
    {
        ExcelOperation mExcelOperation = new ExcelOperation();
        bool bool_ExcelImpoert = mExcelOperation.ExcelImportWithLayoutCheck(this.Application["filePath"].ToString(), "Sheet1");
        if (bool_ExcelImpoert)
        {
            mExcelOperation.TraceFromExcel(DropDownList1.SelectedItem.Value);
            lbl_DateInfo.Text = this.Application["filePath"].ToString().Substring(0, 8);
        }
        this.Page.ClientScript.RegisterStartupScript(this.Page.GetType(), "setECharts2", "setECharts2('" +
            Data2Trace.values_ECharts1_1 + "','"
            + Data2Trace.values_ECharts1_2 + "','"
            + Data2Trace.values_ECharts1_3 + "','"
            + Data2Trace.values_ECharts1_4 + "','"
            + Data2Trace.values_ECharts2_1 + "','"
            + Data2Trace.values_ECharts2_2 + "','"
            + Data2Trace.values_ECharts2_3 + "','"
            + Data2Trace.values_ECharts2_4 + "','"
            + Data2Trace.values_ECharts3_1 + "','"
            + Data2Trace.values_ECharts3_2 + "','"
            + Data2Trace.values_ECharts3_3 + "','"
            + Data2Trace.values_ECharts3_4 + "')", true);
        Label1.Text = this.Application["filePath"].ToString()+DropDownList1.SelectedItem.Value;
    }
    protected void Button1_Click(object sender, EventArgs e)
    {
        ExcelOperation mExcelOperation = new ExcelOperation();
        if (mExcelOperation.colsCount==42)
        {
            bool bool_ExcelImpoert = mExcelOperation.ExcelImportWithLayoutCheck(this.Application["filePath"].ToString(), "Sheet1");
            if (bool_ExcelImpoert)
            {
                mExcelOperation.TraceFromExcel(DropDownList1.SelectedItem.Value);
                lbl_DateInfo.Text = this.Application["filePath"].ToString().Substring(0, 8);
            }
            this.Page.ClientScript.RegisterStartupScript(this.Page.GetType(), "setECharts2", "setECharts2('" +
                Data2Trace.values_ECharts1_1 + "','"
                + Data2Trace.values_ECharts1_2 + "','"
                + Data2Trace.values_ECharts1_3 + "','"
                + Data2Trace.values_ECharts1_4 + "','"
                + Data2Trace.values_ECharts2_1 + "','"
                + Data2Trace.values_ECharts2_2 + "','"
                + Data2Trace.values_ECharts2_3 + "','"
                + Data2Trace.values_ECharts2_4 + "','"
                + Data2Trace.values_ECharts3_1 + "','"
                + Data2Trace.values_ECharts3_2 + "','"
                + Data2Trace.values_ECharts3_3 + "','"
                + Data2Trace.values_ECharts3_4 + "')", true);
            Label1.Text = this.Application["filePath"].ToString() + DropDownList1.SelectedItem.Value;
        }
        
    }
}