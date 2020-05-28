using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;

public partial class _Default : System.Web.UI.Page
{
    ExcelOperation mExcelOperation;
    protected void Page_Load(object sender, EventArgs e)
    {
        mExcelOperation = new ExcelOperation();
        FileTimeInfo mFileTimeInfo = GetLatestFileTimeInfo(@"C:\Users\Public\Music\", ".xlsx");
        this.Application["filePath"] = mFileTimeInfo.FileName;
        lbl_FilePath.Text = mFileTimeInfo.FileName;
    }
    protected void btn_Upload_Click(object sender, EventArgs e)
    {
        string FileName = File1.PostedFile.FileName;
        string DirectoryName = Path.GetFullPath(File1.PostedFile.FileName); //获取文件所在目录
        File1.PostedFile.SaveAs(@"C:\Users\Public\Music\" + FileName);
        this.Application["filePath"] = FileName;
        lbl_FilePath.Text = this.Application["filePath"].ToString();
        
        bool bool_ExcelImpoert = mExcelOperation.ExcelImportWithLayoutCheck(this.Application["filePath"].ToString(), "Sheet1");
        if (bool_ExcelImpoert)//判断导入Excel格式是否符合要求
        {
            mExcelOperation.TraceFromExcel();
        }
    }
    public class FileTimeInfo
    {
        public string FileName;  //文件名
        public DateTime FileCreateTime; //创建时间
    }

    static FileTimeInfo GetLatestFileTimeInfo(string dir, string ext, string type = "CS E2E")
    {
        List<FileTimeInfo> list = new List<FileTimeInfo>();
        DirectoryInfo d = new DirectoryInfo(dir);
        foreach (FileInfo file in d.GetFiles())
        {
            if (file.Extension.ToUpper() == ext.ToUpper() && file.Name.Contains(type))
            {
                list.Add(new FileTimeInfo()
                {
                    FileName = file.Name,
                    FileCreateTime = file.CreationTime
                });
            }
        }
        var f = from x in list
                orderby x.FileCreateTime
                select x;
        return f.LastOrDefault();
    }


    protected void btn_Download_Click(object sender, EventArgs e)
    {
        bool bool_ExcelImpoert = mExcelOperation.ExcelImportWithLayoutCheck(this.Application["filePath"].ToString(), "Sheet1");
        if (bool_ExcelImpoert)
        {
            mExcelOperation.TraceFromExcel();
        }
        
        //For download
        //FileStream fs = new FileStream(@"C:\Users\Public\Documents\" + mExcelOperation.filename_AlertList, FileMode.Open);
        string DownloadFileName = mExcelOperation.filename_edited.Split('_')[0] + "_edited.xlsx";
        FileStream fs = new FileStream(@"C:\Users\Public\Music\" + DownloadFileName, FileMode.Open);
        byte[] file_download = new byte[fs.Length];
        fs.Read(file_download, 0, file_download.Length);
        fs.Close();
        Response.Clear();
        Response.AddHeader("content-disposition", "attachment; filename=" + mExcelOperation.filename_AlertList);//以二进制流模式，强制下载 
        Response.ContentType = "application/octet-stream";
        Response.BinaryWrite(file_download);
    }
    protected void Button1_Click(object sender, EventArgs e)
    {
        //case by case
        string FileName = File2.PostedFile.FileName;
        string DirectoryName = Path.GetFullPath(File2.PostedFile.FileName); //获取文件所在目录
        File2.PostedFile.SaveAs(@"C:\Users\Public\Music\CbyC" + Path.GetExtension(FileName));
        //this.Application["filePath_CbyC"] = FileName;        
    }
    protected void Button2_Click(object sender, EventArgs e)
    {
        //match customer request
        string FileName = File3.PostedFile.FileName;
        string DirectoryName = Path.GetFullPath(File3.PostedFile.FileName); //获取文件所在目录
        File3.PostedFile.SaveAs(@"C:\Users\Public\Music\MatchReq" + Path.GetExtension(FileName));
        //this.Application["filePath_MatchReq"] = FileName;       
    }
}
