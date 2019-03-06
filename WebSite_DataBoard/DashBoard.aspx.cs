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

public partial class DashBoard : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        //this.Page.ClientScript.RegisterStartupScript(this.Page.GetType(), "setLoading", "setLoading()", true);
        //this.Page.ClientScript.RegisterStartupScript(this.Page.GetType(), "setECharts1", "setECharts1('"+this.Application["filePath"].ToString()+"')", true);
        this.Page.ClientScript.RegisterStartupScript(this.Page.GetType(), "setECharts2", "setECharts2('" + this.Application["filePath"].ToString() + "')", true);
        //this.Page.ClientScript.RegisterStartupScript(this.Page.GetType(), "setECharts2", "setECharts2()", true);
    }

    private XSSFSheet sht;
    protected String excelContent;
    protected String excelResult;
    int num_noDIV = 0;
    int num_ok = 0;
    int num_nok = 0;
    int num_ok_noDIV = 0;
    int num_Error_noDIV = 0;
    int num_Warning_noDIV = 0;
    int[] POnum_PerMonth = new int[12];

    private void GetInfoFromXLSX(string filename)
    {
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
            XSSFCell nCell_Create = (XSSFCell)nRow.GetCell(25);
            XSSFCell nCell_Finish = (XSSFCell)nRow.GetCell(28);
            XSSFCell nCell_LT = (XSSFCell)nRow.GetCell(30);
            XSSFCell nCell_Confirmed = (XSSFCell)nRow.GetCell(31);
            DateTimeFormatInfo dtFormat = new DateTimeFormatInfo();
            dtFormat.ShortDatePattern = "yyyy/MM/dd";

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
            dt_SOCreated = Convert.ToDateTime(nCell_Create.DateCellValue);
            if (dt_SOCreated.Year==2018)
            {
                POnum_PerMonth[dt_SOCreated.Month - 1]++;

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

            }
            else
            {
                DateTime dt_Act;
                DateTime dt_Plan;
                try
                {
                    dt_Act = Convert.ToDateTime(nCell_Finish.DateCellValue);
                    dt_Plan = Convert.ToDateTime(nCell_Confirmed.DateCellValue);
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