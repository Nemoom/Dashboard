using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;

public partial class _Default : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
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
        ExcelOperation mExcelOperation = new ExcelOperation();
        bool bool_ExcelImpoert = mExcelOperation.ExcelImportWithLayoutCheck(this.Application["filePath"].ToString(), "Sheet1");
        if (bool_ExcelImpoert)
        { 
        
        }
    }
    public class FileTimeInfo
    {
        public string FileName;  //文件名
        public DateTime FileCreateTime; //创建时间
    }

    static FileTimeInfo GetLatestFileTimeInfo(string dir, string ext)
    {
        List<FileTimeInfo> list = new List<FileTimeInfo>();
        DirectoryInfo d = new DirectoryInfo(dir);
        foreach (FileInfo file in d.GetFiles())
        {
            if (file.Extension.ToUpper() == ext.ToUpper())
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


}
