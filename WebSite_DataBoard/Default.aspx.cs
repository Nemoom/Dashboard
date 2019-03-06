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
        lbl_FilePath.Text = this.Application["filePath"].ToString();
    }
    protected void btn_Upload_Click(object sender, EventArgs e)
    {
        string FileName = File1.PostedFile.FileName;
        string DirectoryName = Path.GetFullPath(File1.PostedFile.FileName); //获取文件所在目录
        File1.PostedFile.SaveAs(@"C:\Users\Public\Music\" + FileName);
        this.Application["filePath"] = FileName;
        lbl_FilePath.Text = this.Application["filePath"].ToString();
    }
}
