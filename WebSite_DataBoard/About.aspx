<%@ Page Title="About Us" Language="C#" MasterPageFile="~/Site.master" AutoEventWireup="true"
    CodeFile="About.aspx.cs" Inherits="About" %>

<asp:Content ID="HeaderContent" runat="server" ContentPlaceHolderID="HeadContent">
</asp:Content>
<asp:Content ID="BodyContent" runat="server" ContentPlaceHolderID="MainContent">
    <h2>
        About
    </h2>
    <p>
        <asp:Label ID="list" runat="server"></asp:Label>
    </p>
    <script type="text/javascript">
        var x = window.screen.height;
        var y = window.screen.width;
        var msg = "您屏幕分辨率：" + x + "x" + y + "像素";
        alert(msg);
    </script>
</asp:Content>

