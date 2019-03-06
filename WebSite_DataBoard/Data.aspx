<%@ Page Title="" Language="C#" MasterPageFile="~/Site.master" AutoEventWireup="true" CodeFile="Data.aspx.cs" Inherits="Data" %>

<asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" Runat="Server">
    <style type="text/css">
        .style2
        {
            width: 297px;
        }
        .style3
        {
            width: 339px;
        }
    </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" Runat="Server">    
    <table style="width: 100%;">
        <tr>
            <td class="style3"> 
                <asp:CheckBox ID="CBox_Status" runat="server" />              
                <asp:Label ID="Label1" runat="server" Text="Status" Width="120px"></asp:Label>
                <asp:DropDownList ID="DList_Status" Width="180px" runat="server" >
                    
                </asp:DropDownList>
            </td>
            <td class="style2">
                <asp:CheckBox ID="CBox_Party" runat="server" />
                <asp:Label ID="Label2" runat="server" Text="Sold-To Party" Width="120px"></asp:Label>
                <asp:DropDownList ID="DList_Party" Width="180px" runat="server">
                    
                </asp:DropDownList>
            </td>
            <td>
                <asp:CheckBox ID="CBox_Segment" runat="server" />
                <asp:Label ID="Label3" runat="server" Text="Segment" Width="120px"></asp:Label>
                <asp:DropDownList ID="DList_Segment" Width="180px" runat="server">
                    
                </asp:DropDownList>
            </td>
        </tr>
        <tr>
            <td class="style3">
                <asp:CheckBox ID="CBox_Sales" runat="server" />
                <asp:Label ID="Label4" runat="server" Text="Sales" Width="120px"></asp:Label>
                <asp:DropDownList ID="DList_Sales" Width="180px" runat="server">
                    
                </asp:DropDownList>
            </td>
            <td class="style2">
                <asp:CheckBox ID="CBox_SOffice" runat="server" />
                <asp:Label ID="Label5" runat="server" Text="Sales Office" Width="120px"></asp:Label>
                <asp:DropDownList ID="DList_SOffice" Width="180px" runat="server">
                    
                </asp:DropDownList>
            </td>
            <td>
                <asp:CheckBox ID="CBox_Engineer" runat="server" />
                <asp:Label ID="Label6" runat="server" Text="Engineer" Width="120px"></asp:Label>
                <asp:DropDownList ID="DList_Engineer" Width="180px" runat="server">
                    
                </asp:DropDownList>
            </td>
        </tr>
        <tr>
            <td class="style3">
                <asp:CheckBox ID="CBox_FieldName" runat="server" />
                <asp:Label ID="Label7" runat="server" Text="Field Name" Width="120px"></asp:Label>
                <asp:DropDownList ID="DList_FieldName" Width="180px" runat="server"></asp:DropDownList>                
            </td>
            <td class="style2">
                <asp:Label ID="Label8" runat="server" Text="Keyword" Width="120px"></asp:Label>
                <asp:TextBox ID="txt_Keyword" runat="server"></asp:TextBox>
            </td>
            <td align="right" style="padding: 0px 17px 0px 0px; margin: 0px 10px 0px 10px;">
                <asp:Button ID="btn_Search" runat="server" Text="Search" 
                    onclick="btn_Search_Click" />
                <asp:Button ID="btn_Export" runat="server" Text="Export" 
                    onclick="btn_Export_Click"/>
            </td>
        </tr>
    </table>
    <div>
        <asp:Button ID="btn_NewOrder" runat="server" Text="每日新订单" width="16%" 
            onclick="btn_NewOrder_Click"/>
        <asp:Button ID="btn_NDelivery" runat="server" Text="未释放订单" width="16%" 
            onclick="btn_NDelivery_Click"/>
        <asp:Button ID="btn_AbnormalOrder" runat="server" Text="异常订单" width="16%" 
            onclick="btn_AbnormalOrder_Click"/>
        <asp:Button ID="btn_CheckRemind" runat="server" Text="物料检查提醒" width="16%" 
            onclick="btn_CheckRemind_Click"/>
        <asp:Button ID="btn_PlanRemind" runat="server" Text="生产计划提醒" width="16%" 
            onclick="btn_PlanRemind_Click"/>
        <asp:Button ID="btn_DelayOrder" runat="server" Text="延期订单" width="16%" 
            onclick="btn_DelayOrder_Click"/>
    </div>
    <div id="GridViewDiv" style="height: 500px;width:100%; overflow: scroll;text-align:center;" runat="server">
        <asp:GridView ID="GridView1" runat="server" BackColor="White" 
            BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" CellPadding="3" 
            Width="95%">
            <FooterStyle BackColor="White" ForeColor="#000066" />
            <HeaderStyle BackColor="#006699" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="White" ForeColor="#000066" HorizontalAlign="Left" />
            <RowStyle ForeColor="#000066" />
            <SelectedRowStyle BackColor="#669999" Font-Bold="True" ForeColor="White" />
            <SortedAscendingCellStyle BackColor="#F1F1F1" />
            <SortedAscendingHeaderStyle BackColor="#007DBB" />
            <SortedDescendingCellStyle BackColor="#CAC9C9" />
            <SortedDescendingHeaderStyle BackColor="#00547E" />
        </asp:GridView>
    </div>
</asp:Content>

