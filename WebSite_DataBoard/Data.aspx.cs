using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using GlobalParas;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using NPOI.SS.Util;
using NPOI.Util;
using NPOI.XSSF.UserModel;

public partial class Data : System.Web.UI.Page
{
    static OleDbConnection conn_Cur = null;
    protected void Page_Load(object sender, EventArgs e)
    {
        GlobalParas.GlobalParas.mDataSet = ExcelToDataSet(this.Application["filePath"].ToString(), "Sheet1");
        string[] arrayFieldName = new string[GlobalParas.GlobalParas.mDataSet.Tables[0].Columns.Count];
        for (int i = 0; i < GlobalParas.GlobalParas.mDataSet.Tables[0].Columns.Count; i++)
        {
            arrayFieldName[i] = GlobalParas.GlobalParas.mDataSet.Tables[0].Columns[i].ToString();
        }
        DList_FieldName.DataSource = arrayFieldName;
        DList_FieldName.DataBind();
        //GridView1.DataSource = GlobalParas.GlobalParas.mDataSet;
        //GridView1.DataBind();

        DataTable dt_Segment = GlobalParas.GlobalParas.mDataSet.Tables[0].DefaultView.ToTable(true, "Segment");
        DList_Segment.DataSource = ItemList2StringArray(dt_Segment);
        DList_Segment.DataBind();
        DataTable dt_Party = GlobalParas.GlobalParas.mDataSet.Tables[0].DefaultView.ToTable(true, "Sold-To Party");
        DList_Party.DataSource = ItemList2StringArray(dt_Party);
        DList_Party.DataBind();
        DataTable dt_Engineer = GlobalParas.GlobalParas.mDataSet.Tables[0].DefaultView.ToTable(true, "Proj Engineer Name");
        DList_Engineer.DataSource = ItemList2StringArray(dt_Engineer);
        DList_Engineer.DataBind();
        DataTable dt_SOffice = GlobalParas.GlobalParas.mDataSet.Tables[0].DefaultView.ToTable(true, "Sales Office");
        DList_SOffice.DataSource = ItemList2StringArray(dt_SOffice);
        DList_SOffice.DataBind();
        DataTable dt_Sales = GlobalParas.GlobalParas.mDataSet.Tables[0].DefaultView.ToTable(true, "Sales Responsible");
        DList_Sales.DataSource = ItemList2StringArray(dt_Sales);
        DList_Sales.DataBind();
        DataTable dt_Status = GlobalParas.GlobalParas.mDataSet.Tables[0].DefaultView.ToTable(true, "Status");
        DList_Status.DataSource = ItemList2StringArray(dt_Status);
        DList_Status.DataBind();
        GlobalParas.GlobalParas.mDataSet.Dispose();
        this.GridViewDiv.Visible = true;
    }
    public string[] ItemList2StringArray(DataTable mItemList)
    {
        string[] arrayString = new string[mItemList.Rows.Count];
        for (int i = 0; i < mItemList.Rows.Count; i++)
        {
            arrayString[i] = (string)mItemList.Rows[i].ItemArray[0];
        }
        return arrayString;
    }
    public void ReadMyData(string connectionString)
    {
        string queryString = "SELECT OrderID, CustomerID FROM Orders";
        using (OleDbConnection connection = new OleDbConnection(connectionString))
        {
            OleDbCommand command = new OleDbCommand(queryString, connection);
            connection.Open();
            OleDbDataReader reader = command.ExecuteReader();

            while (reader.Read())
            {
                Console.WriteLine(reader.GetInt32(0) + ", " + reader.GetString(1));
            }
            // always call Close when done reading.
            reader.Close();
        }
    }
    protected void btn_Search_Click(object sender, EventArgs e)
    {
        string sqlStr = "";

        if (CBox_Status.Checked)
        {
            if (sqlStr=="")
            {
                sqlStr = " where [Status]='" + DList_Status.SelectedValue.ToString() + "'";
            }
            else
            {
                sqlStr = sqlStr + " and [Status]='" + DList_Status.SelectedValue.ToString() + "'";
            }
        }
        if (CBox_Party.Checked)
        {
            if (sqlStr == "")
            {
                sqlStr = " where [Sold-To Party]='" + DList_Party.SelectedValue.ToString() + "'";
            }
            else
            {
                sqlStr = sqlStr + " and [Sold-To Party]='" + DList_Party.SelectedValue.ToString() + "'";
            }
        }
        GlobalParas.GlobalParas.mDataSet = SelectCMD(sqlStr, @"C:\Users\Public\Music\" + this.Application["filePath"].ToString());
        GridView1.DataSource = GlobalParas.GlobalParas.mDataSet;
        GridView1.DataBind();

    }
    public bool Excel2DataSet(string FilePath)
    {
        try
        {
            XSSFWorkbook wb = new XSSFWorkbook(File.OpenRead(this.Application["filePath"].ToString()));
            XSSFSheet sht = (XSSFSheet)wb.GetSheet("Sheet1");
            DataTable table=new DataTable();
            GlobalParas.GlobalParas.mDataSet.Tables.Add(table);

            return true;
        }
        catch (Exception)
        {
            return false;            
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
        if (string.IsNullOrEmpty(sheetName)) { sheetName = GetExcelFirstTableName(pathName); }
        if (!File.Exists(pathName)) { throw new Exception("文件不存在"); }
        
        //读取Excel里面有 表Sheet1
        OleDbDataAdapter oda = new OleDbDataAdapter(string.Format("select * from [{0}$]", sheetName), GetConnection(pathName));
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
            //if (conn_Cur==null)
            //{
            //    conn_Cur = GetConnection(pathName);
            //}
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
        if (string.IsNullOrEmpty(sheetName)) { sheetName = "Sheet1"; }
        OleDbConnection cnnxls = GetConnection(pathName);
        OleDbDataAdapter oda = new OleDbDataAdapter(string.Format("select * from [{0}$]", sheetName) + sqlStr, cnnxls);
        //将Excel里面有表内容装载到内存表中！
        oda.Fill(ds);
        oda.Dispose();
        cnnxls.Close();
        return ds;
    }
    public void CallLoadingPage()
    {
        FileStream stream = new FileStream(Server.MapPath("Loading_icon.gif"), FileMode.Open);
        long FileSize = stream.Length;
        byte[] Buffer = new byte[(int)FileSize];
        stream.Read(Buffer, 0, (int)FileSize);
        stream.Close();
        Response.BinaryWrite(Buffer);
    }
    protected void btn_NewOrder_Click(object sender, EventArgs e)
    {
        string sqlStr = " where [PrdCreatedOn]=#" + DateTime.Today.Year + "-" + DateTime.Today.Month + "-" + DateTime.Today.Day+"#";
        GlobalParas.GlobalParas.mDataSet = SelectCMD(" where [PrdCreatedOn]=#2018-4-26#", @"C:\Users\Public\Music\" + this.Application["filePath"].ToString());
        //GlobalParas.GlobalParas.mDataSet = SelectCMD(sqlStr, @"C:\Users\Public\Music\" + this.Application["filePath"].ToString());
        GridView1.DataSource = GlobalParas.GlobalParas.mDataSet;
        GridView1.DataBind();
    }
    protected void btn_NDelivery_Click(object sender, EventArgs e)
    {
        string sqlStr = " where ";
        //GlobalParas.GlobalParas.mDataSet = SelectCMD(" where [Actual Release Date] is null", @"C:\Users\Public\Music\" + this.Application["filePath"].ToString());
        GlobalParas.GlobalParas.mDataSet = SelectCMD(" where [Actual Release Date] is null and [PrdCreatedOn]<#2018-10-10#", @"C:\Users\Public\Music\" + this.Application["filePath"].ToString());
        GridView1.DataSource = GlobalParas.GlobalParas.mDataSet;
        GridView1.DataBind();
    }
    protected void btn_AbnormalOrder_Click(object sender, EventArgs e)
    {
        string sqlStr = " where ";
        GlobalParas.GlobalParas.mDataSet = SelectCMD(" where [Segment]='PLAS'", @"C:\Users\Public\Music\" + this.Application["filePath"].ToString());
        GridView1.DataSource = GlobalParas.GlobalParas.mDataSet;
        GridView1.DataBind();
    }
    protected void btn_CheckRemind_Click(object sender, EventArgs e)
    {
        DateTime startDay = DateTime.Today.AddDays(7);
        DateTime endDay = DateTime.Today.AddDays(14);
        string sqlStr = " where [1st ConfirmedDt] between #" + startDay.Year + "-" + startDay.Month + "-" + startDay.Day
            + "# and #" + endDay.Year + "-" + endDay.Month + "-" + endDay.Day + "#";
        GlobalParas.GlobalParas.mDataSet = SelectCMD(sqlStr, @"C:\Users\Public\Music\" + this.Application["filePath"].ToString());
        GridView1.DataSource = GlobalParas.GlobalParas.mDataSet;
        GridView1.DataBind();

    }
    protected void btn_PlanRemind_Click(object sender, EventArgs e)
    {
        DateTime startDay = DateTime.Today;
        DateTime endDay = DateTime.Today.AddDays(7);
        string sqlStr = " where [1st ConfirmedDt] between #" + startDay.Year + "-" + startDay.Month + "-" + startDay.Day
            + "# and #" + endDay.Year + "-" + endDay.Month + "-" + endDay.Day + "#";
        GlobalParas.GlobalParas.mDataSet = SelectCMD(sqlStr, @"C:\Users\Public\Music\" + this.Application["filePath"].ToString());
        GridView1.DataSource = GlobalParas.GlobalParas.mDataSet;
        GridView1.DataBind();
        if (GridView1.Columns.Count != 0)
        {
            for (int i = 10; i < 24; i++)
            {
                GridView1.Columns[i].Visible = false;
            }
        }
              
    }
    protected void btn_DelayOrder_Click(object sender, EventArgs e)
    {
        //[SO CreatedOn]起始时间 [Quotation LT]承诺多少工作日完成
        //判断创建日期是周几，进而得出结束时间
    }
    protected void btn_Export_Click(object sender, EventArgs e)
    {

    }
}