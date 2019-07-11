using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.IO;
using System.Data.OleDb;

public partial class Details : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        string loadInfo = Request.Params["Info"];
        string[] str_loadInfo = loadInfo.Split(';');
        string str_chartIndex, str_paramsName, str_seriesIndex, str_dataIndex, str_paramsValue;
        str_chartIndex = str_loadInfo[0];
        str_paramsName = str_loadInfo[1];
        str_seriesIndex = str_loadInfo[2];
        str_dataIndex = str_loadInfo[3];
        str_paramsValue = str_loadInfo[4];
        string sqlStr = "";
        DataSet mDataSet_Detail;
        if (this.Application["filePath"].ToString().Split('_').Length==2)
        {
            switch (str_chartIndex)
            {
                case "1_1":
                    //DCR
                    break;
                case "1_2":
                    //LT
                    break;
                case "1_3":
                    //DCR ALERT
                    switch (str_paramsName)
                    {
                        case "Material Planning Alert":
                            break;
                        case "Production Planning Alert":
                            break;
                        case "DC Failed and Match Req":
                            sqlStr = " where [Real Failed]='DC Failed'";
                            break;
                        case "DC Failed and Req Failed":
                            sqlStr = " where [Real Failed]='Real Failed'";
                            break;
                        default:
                            break;
                    }
                    mDataSet_Detail = SelectCMD(sqlStr, @"C:\Users\Public\Music\" + this.Application["filePath"].ToString());
                    GridView1.DataSource = mDataSet_Detail;
                    GridView1.DataBind();
                    break;
                case "1_4":
                    //Ongoing_LT
                    break;
                case "2_1":
                    //SO Income
                    break;
                case "2_2":
                    //Finished PO
                    break;
                case "2_3":
                    //Shippment
                    break;
                case "2_4":
                    //Non-delivery
                    break;
                case "3_1":
                    //0400 SO->PO
                    sqlStr = " where [DLV Plant]='0400'";
                    //str_seriesIndex 总创建/延期创建
                    if (str_seriesIndex == "1")
                    {
                        //延期创建
                        sqlStr = sqlStr + " and [SO -> PO Creation(WD)]>1";
                    }
                    //str_dataIndex 哪个月创建的PO
                    if (sqlStr == "")
                    {
                        sqlStr = " where [Prod CreatedOn] between #" + DateTime.Today.Year + "-" + (Convert.ToInt16(str_dataIndex) + 1).ToString() + "-1"
                                + "# and #" + DateTime.Today.Year + "-" + (Convert.ToInt16(str_dataIndex) + 1).ToString() + "-" + DateTime.DaysInMonth(DateTime.Today.Year, Convert.ToInt16(str_dataIndex) + 1).ToString() + "#";
                    }
                    else
                    {
                        sqlStr = sqlStr + " and [Prod CreatedOn] between #" + DateTime.Today.Year + "-" + (Convert.ToInt16(str_dataIndex) + 1).ToString() + "-1"
                                + "# and #" + DateTime.Today.Year + "-" + (Convert.ToInt16(str_dataIndex) + 1).ToString() + "-" + DateTime.DaysInMonth(DateTime.Today.Year, Convert.ToInt16(str_dataIndex) + 1).ToString() + "#";
                    }
                    mDataSet_Detail = SelectCMD(sqlStr, @"C:\Users\Public\Music\" + this.Application["filePath"].ToString());
                    if (mDataSet_Detail.Tables[0].Rows.Count == Convert.ToInt16(str_paramsValue))
                    {
                        GridView1.DataSource = mDataSet_Detail;
                        GridView1.DataBind();
                        //Response.Write("<script>alert('OK')</script>");
                    }
                    else
                    {
                        Response.Write("<script>alert('前后查询数量不匹配')</script>");
                    }
                    break;
                case "3_2":
                    //0400 PO->PO Release
                    sqlStr = " where [DLV Plant]='0400'";
                    //str_seriesIndex 总创建/延期创建
                    if (str_seriesIndex == "1")
                    {
                        //延期创建
                        sqlStr = sqlStr + " and [PO Creation -> PO Release(WD)]>5";
                    }
                    //str_dataIndex 哪个月创建的PO
                    if (sqlStr == "")
                    {
                        sqlStr = " where [Actual Release Date] between #" + DateTime.Today.Year + "-" + (Convert.ToInt16(str_dataIndex) + 1).ToString() + "-1"
                                + "# and #" + DateTime.Today.Year + "-" + (Convert.ToInt16(str_dataIndex) + 1).ToString() + "-" + DateTime.DaysInMonth(DateTime.Today.Year, Convert.ToInt16(str_dataIndex) + 1).ToString() + "#";
                    }
                    else
                    {
                        sqlStr = sqlStr + " and [Actual Release Date] between #" + DateTime.Today.Year + "-" + (Convert.ToInt16(str_dataIndex) + 1).ToString() + "-1"
                                + "# and #" + DateTime.Today.Year + "-" + (Convert.ToInt16(str_dataIndex) + 1).ToString() + "-" + DateTime.DaysInMonth(DateTime.Today.Year, Convert.ToInt16(str_dataIndex) + 1).ToString() + "#";
                    }
                    mDataSet_Detail = SelectCMD(sqlStr, @"C:\Users\Public\Music\" + this.Application["filePath"].ToString());
                    if (mDataSet_Detail.Tables[0].Rows.Count == Convert.ToInt16(str_paramsValue))
                    {
                        GridView1.DataSource = mDataSet_Detail;
                        GridView1.DataBind();
                        //Response.Write("<script>alert('OK')</script>");
                    }
                    else
                    {
                        Response.Write("<script>alert('前后查询数量不匹配')</script>");
                    }
                    break;
                case "3_3":
                    //PO Release->Actual Finish
                    //延期完成
                    sqlStr = " where [PO Release -> Actual Finish(WD)]>23";

                    //str_seriesIndex 0400/0481/0400+0481
                    switch (str_seriesIndex)
                    {
                        case "0":
                            sqlStr = sqlStr + " and [DLV Plant]='0400'";
                            break;
                        case "1":
                            sqlStr = sqlStr + " and [DLV Plant]='0481'";
                            break;
                        case "2":
                            break;
                        default:
                            break;
                    }                  
                   
                    //str_dataIndex 哪个月创建的PO
                    if (sqlStr == "")
                    {
                        sqlStr = " where [Actual Finish Date] between #" + DateTime.Today.Year + "-" + (Convert.ToInt16(str_dataIndex) + 1).ToString() + "-1"
                                + "# and #" + DateTime.Today.Year + "-" + (Convert.ToInt16(str_dataIndex) + 1).ToString() + "-" + DateTime.DaysInMonth(DateTime.Today.Year, Convert.ToInt16(str_dataIndex) + 1).ToString() + "#";
                    }
                    else
                    {
                        sqlStr = sqlStr + " and [Actual Finish Date] between #" + DateTime.Today.Year + "-" + (Convert.ToInt16(str_dataIndex) + 1).ToString() + "-1"
                                + "# and #" + DateTime.Today.Year + "-" + (Convert.ToInt16(str_dataIndex) + 1).ToString() + "-" + DateTime.DaysInMonth(DateTime.Today.Year, Convert.ToInt16(str_dataIndex) + 1).ToString() + "#";
                    }
                    mDataSet_Detail = SelectCMD(sqlStr, @"C:\Users\Public\Music\" + this.Application["filePath"].ToString());

                    GridView1.DataSource = mDataSet_Detail;
                    GridView1.DataBind();

                    break;
                case "3_4":
                    //Actual Finish->Ex-plant
                    //延期完成
                    sqlStr = " where [Actual Finish -> Ex-plant(WD)]>2";

                    //str_seriesIndex 0400/0481/0400+0481
                    switch (str_seriesIndex)
                    {
                        case "0":
                            sqlStr = sqlStr + " and [DLV Plant]='0400'";
                            break;
                        case "1":
                            sqlStr = sqlStr + " and [DLV Plant]='0481'";
                            break;
                        case "2":
                            break;
                        default:
                            break;
                    }

                    //str_dataIndex 哪个月创建的PO
                    if (sqlStr == "")
                    {
                        sqlStr = " where [ShipmentStartOn] between #" + DateTime.Today.Year + "-" + (Convert.ToInt16(str_dataIndex) + 1).ToString() + "-1"
                                + "# and #" + DateTime.Today.Year + "-" + (Convert.ToInt16(str_dataIndex) + 1).ToString() + "-" + DateTime.DaysInMonth(DateTime.Today.Year, Convert.ToInt16(str_dataIndex) + 1).ToString() + "#";
                    }
                    else
                    {
                        sqlStr = sqlStr + " and [ShipmentStartOn] between #" + DateTime.Today.Year + "-" + (Convert.ToInt16(str_dataIndex) + 1).ToString() + "-1"
                                + "# and #" + DateTime.Today.Year + "-" + (Convert.ToInt16(str_dataIndex) + 1).ToString() + "-" + DateTime.DaysInMonth(DateTime.Today.Year, Convert.ToInt16(str_dataIndex) + 1).ToString() + "#";
                    }
                    mDataSet_Detail = SelectCMD(sqlStr, @"C:\Users\Public\Music\" + this.Application["filePath"].ToString());

                    GridView1.DataSource = mDataSet_Detail;
                    GridView1.DataBind();
                    break;
                default:
                    break;
            }
        }//文件名中有_edited
        else
        {
            //文件名中没有_edited，无法拉出详单
        }
        
    }

    /// <summary>
    /// 把数据从Excel装载到DataSet
    /// </summary>
    /// <param name="pathName">带路径的Excel文件名</param>
    /// <param name="sheetName">工作表名</param>
    /// <param name="tbContainer">将数据存入的DataSet</param>
    /// <returns></returns>
    public DataSet ExcelToDataSet(string pathName, string sheetName = null)
    {
        DataSet ds = new DataSet();
        string strConn = string.Empty;
        pathName = @"C:\Users\Public\Music\" + pathName;
        if (string.IsNullOrEmpty(sheetName))
        {
            sheetName = GetExcelFirstTableName(pathName);
        }
        else if (!sheetName.EndsWith("$"))
        {
            sheetName = sheetName + "$";
        }
        if (!File.Exists(pathName)) { throw new Exception("文件不存在"); }

        //读取Excel里面有 表Sheet1
        OleDbDataAdapter oda = new OleDbDataAdapter(string.Format("select * from [{0}]", sheetName), GetConnection(pathName));
        //将Excel里面有表内容装载到内存表中！
        oda.Fill(ds);
        oda.Dispose();
        return ds;
    }

    /// <summary>
    /// C#中获取Excel文件的第一个表名
    /// Excel文件中第一个表名的缺省值是Sheet1$, 但有时也会被改变为其他名字. 如果需要在C#中使用OleDb读写Excel文件, 就需要知道这个名字是什么. 以下代码就是实现这个功能的:
    /// </summary>
    /// <param name="pathName"></param>
    /// <returns></returns>
    public static string GetExcelFirstTableName(string pathName)
    {
        //string pathName = @"C:\Users\Public\Music\" + excelFileName;
        string tableName = null;
        if (File.Exists(pathName))
        {
            OleDbConnection conn_Cur = null;
            using (conn_Cur = GetConnection(pathName))
            {
                conn_Cur.Open();
                DataTable dt = conn_Cur.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                tableName = dt.Rows[0][2].ToString().Trim();
            }
        }
        return tableName;
    }

    /// <summary>
    /// 利用excelFileName生成OleDbConnection，可用来赋值全局变量OleDbConnection
    /// </summary>
    /// <param name="pathName">完整的文件路径</param>
    /// <returns></returns>
    public static OleDbConnection GetConnection(string pathName)
    {
        string strConn = string.Empty;
        //string pathName = @"C:\Users\Public\Music\" + excelFileName;
        FileInfo file = new FileInfo(pathName);
        if (!file.Exists) { throw new Exception("文件不存在"); }
        string extension = file.Extension.ToLower();
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
        return cnnxls;
    }

    public DataSet SelectCMD(string sqlStr, string pathName, string sheetName = null)
    {
        DataSet ds = new DataSet();
        if (string.IsNullOrEmpty(sheetName))
        {
            sheetName = GetExcelFirstTableName(pathName);
        }
        else if (!sheetName.EndsWith("$"))
        {
            sheetName = sheetName + "$";
        }
        OleDbConnection cnnxls = GetConnection(pathName);
        OleDbDataAdapter oda = new OleDbDataAdapter(string.Format("select * from [{0}]", sheetName) + sqlStr, cnnxls);
        //将Excel里面有表内容装载到内存表中！
        oda.Fill(ds);
        oda.Dispose();
        cnnxls.Close();
        return ds;
    }
}