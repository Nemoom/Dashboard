<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.master" AutoEventWireup="true"
    CodeFile="Default.aspx.cs" Inherits="_Default" %>

<asp:Content ID="HeaderContent" runat="server" ContentPlaceHolderID="HeadContent">
</asp:Content>
<asp:Content ID="BodyContent" runat="server" ContentPlaceHolderID="MainContent">
    <h2>
        Welcome to E2E Dashboard!
    </h2>
    <h3>Current File Path:</h3><dfn><asp:Label ID="lbl_FilePath" runat="server" Text=""></asp:Label></dfn>
    <h3>Load New File: </h3>
    <input id="File1" type="file"  runat="server"/>
    <asp:Button ID="btn_Upload" runat="server" Text="Upload" 
        onclick="btn_Upload_Click"/>
    <asp:Button ID="btn_Download" runat="server" Text="Download" 
        onclick="btn_Download_Click" />
</asp:Content>
