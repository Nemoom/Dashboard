using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class About : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        HttpBrowserCapabilities bc = Request.Browser;

        list.Text = "";
        list.Text += "操作系统：" + bc.Platform + "<br>";
        list.Text += "是否是 Win16 系统：" + bc.Win16 + "<br>";
        list.Text += "是否是 Win32 系统：" + bc.Win32 + "<br>";
        list.Text += "---<br>";

        list.Text += "浏览器：" + bc.Browser + "<br>";
        list.Text += "浏览器标识：" + bc.Id + "<br>";
        list.Text += "浏览器版本：" + bc.Version + "<br>";
        list.Text += "浏览器 MajorVersion：" + bc.MajorVersion.ToString() + "<br>";
        list.Text += "浏览器 MinorVersion：" + bc.MinorVersion.ToString() + "<br>";
        list.Text += "浏览器是否是测试版本：" + bc.Beta.ToString() + "<br>";
        list.Text += "是否是 America Online 浏览器：" + bc.AOL + "<br>";
        list.Text += "客户端安装的 .NET Framework 版本：" + bc.ClrVersion + "<br>"; //即使安装了 .NET Framework，如果不是 IE 浏览器，检测版本都是 0.0。
        list.Text += "是否是搜索引擎的网络爬虫：" + bc.Crawler + "<br>";
        list.Text += "是否是移动设备：" + bc.IsMobileDevice + "<br>";
        list.Text += "---<br>";

        list.Text += "显示的颜色深度：" + bc.ScreenBitDepth + "<br>";
        list.Text += "显示的近似宽度（以字符行为单位）：" + bc.ScreenCharactersWidth + "<br>";
        list.Text += "显示的近似高度（以字符行为单位）：" + bc.ScreenCharactersHeight + "<br>";
        list.Text += "显示的近似宽度（以像素行为单位）：" + bc.ScreenPixelsWidth + "<br>";
        list.Text += "显示的近似高度（以像素行为单位）：" + bc.ScreenPixelsHeight + "<br>";
        list.Text += "---<br>";

        list.Text += "是否支持 CSS：" + bc.SupportsCss + "<br>";
        list.Text += "是否支持 ActiveX 控件：" + bc.ActiveXControls.ToString() + "<br>";
        list.Text += "是否支持 JavaApplets：" + bc.JavaApplets.ToString() + "<br>";
        list.Text += "是否支持 JavaScript：" + bc.JavaScript.ToString() + "<br>";
        list.Text += "JScriptVersion：" + bc.JScriptVersion.ToString() + "<br>";
        list.Text += "是否支持 VBScript：" + bc.VBScript.ToString() + "<br>";
        list.Text += "是否支持 Cookies：" + bc.Cookies + "<br>";
        list.Text += "支持的 MSHTML 的 DOM 版本：" + bc.MSDomVersion + "<br>";
        list.Text += "支持的 W3C 的 DOM 版本：" + bc.W3CDomVersion + "<br>";
        list.Text += "是否支持通过 HTTP 接收 XML：" + bc.SupportsXmlHttp + "<br>";
        list.Text += "是否支持框架：" + bc.Frames.ToString() + "<br>";
        list.Text += "超链接 a 属性 href 值的最大长度：" + bc.MaximumHrefLength + "<br>";
        list.Text += "是否支持表格：" + bc.Tables + "<br>";
    }
}
