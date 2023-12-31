﻿<%@ Page Language="VB" MasterPageFile="~/Site.master" AutoEventWireup="false" CodeFile="MyReportsAdmin.aspx.vb" Inherits="MyReportsAdmin" Title="My Reports Admin" %>








<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="Server">







<asp:UpdatePanel ID="UpdatePanel1" runat="server">
<ContentTemplate>

<asp:Label ID="lblDatabase" runat="server"  style="font-weight:bold; font-size:1.4em; color:Red;"></asp:Label>
<asp:Panel ID="pnlUpdateStatus" runat="server" Visible="false" CssClass="LunchRotaPanel">
<table  style="border-collapse:collapse; padding:10px; "  >
<tr>
<td colspan="2"  class="EditPanelTitlesTableCell">
<asp:Label ID="lblHeader" runat="server"  style="font-weight:bold; font-size:1.4em"></asp:Label>
<br /><br />
</td>
</tr>
<tr>
<td style="width: 300px;">
      <asp:GridView ID="gv1" runat="server" 
        DataKeyNames="UserFK"
        AutoGenerateColumns="False" 
        AllowPaging="false" 
        pagesize="40"
        AllowSorting="false" 
        AutoGenerateEditButton="false" 
        AutoGenerateDeleteButton="false"
        OnSelectedIndexChanged="gv1_SelectedIndexChanged" 
        OnDataBound="gv1_DataBound" 
        OnPageIndexChanging="gv1_PageIndexChanging"
        OnRowUpdated="gv1_RowUpdated"           
        OnSorting="gv1_Sorting" CssClass="GridviewLunchRota">
          <Columns>
              <asp:BoundField DataField="UserFK" ReadOnly="True" ControlStyle-CssClass="hideGridColumn" ItemStyle-CssClass="hideGridColumn" HeaderStyle-CssClass="hideGridColumn" />
              <asp:BoundField DataField="User_Login" ReadOnly="True"  ControlStyle-CssClass="hideGridColumn" ItemStyle-CssClass="hideGridColumn" HeaderStyle-CssClass="hideGridColumn" />            
              <asp:BoundField DataField="Phone_VeriCode" ReadOnly="True"  ControlStyle-CssClass="hideGridColumn" ItemStyle-CssClass="hideGridColumn" HeaderStyle-CssClass="hideGridColumn"/>
              <asp:BoundField DataField="Email_VeriCode" ReadOnly="True"  ControlStyle-CssClass="hideGridColumn" ItemStyle-CssClass="hideGridColumn" HeaderStyle-CssClass="hideGridColumn" />
              <asp:BoundField DataField="Employee" HeaderText="Employee" ReadOnly="True" ItemStyle-Width="100"  ItemStyle-Wrap="False" />
              <asp:TemplateField HeaderText="Phone" SortExpression="Phone_Number" >
                    <EditItemTemplate>
                    <asp:DropDownList ID="ddlstEditInternationalCodes" runat="server">
                    </asp:DropDownList>
                    <asp:TextBox ID="txtPhone_Number" runat="server" Text='<%# Bind("Phone_Number") %>'></asp:TextBox><br />
                    <asp:Label ID="lblPhoneError" runat="server" forecolor="red"></asp:Label>
                    </EditItemTemplate>
                  <ItemTemplate>
                      <%#Eval("Phone_Number")%>
                  </ItemTemplate>
                  <HeaderStyle ForeColor="Blue" />
                  <ItemStyle HorizontalAlign="Left" VerticalAlign="Top" />
              </asp:TemplateField>

              <asp:TemplateField HeaderText="Email" SortExpression="CompanyFK">
                  <EditItemTemplate>
                      <asp:TextBox ID="txtEmail_Address" runat="server" Text='<%# Bind("Email_Address") %>'></asp:TextBox><br />
                      <asp:Label ID="lblEmailError" runat="server" forecolor="red"></asp:Label>
                  </EditItemTemplate>
                  <ItemTemplate>
                      <%#Eval("Email_Address")%>
                  </ItemTemplate>
                  <HeaderStyle ForeColor="Blue" />
                  <ItemStyle HorizontalAlign="Left" VerticalAlign="Top" />
              </asp:TemplateField>
              <asp:CommandField ShowEditButton="True" />
              
          </Columns>
      </asp:GridView>



</td>

</tr>
<tr>
<td >
<asp:Label ID="lblErrors" runat="server" forecolor="red" Font-Bold="true" Font-Size="Large"></asp:Label>
</td>
</tr>
<tr>
<td >
    <p style="text-align: right">

    </p>
</td>
</tr>
</table>




<div style="display:none">

<asp:Button ID="cmdGetData" runat="server" Text="Go" Visible="false" />
<asp:Label ID="lblUserIP" runat="server" visble="false"></asp:Label>
<asp:Label ID="lblUserID" runat="server" visble="false"></asp:Label>
<asp:Label ID="lblGVCellEmailError" runat="server" visble="false"></asp:Label>
<asp:Label ID="lblGVCellPhoneError" runat="server" visble="false"></asp:Label>
</div>



</asp:Panel>



</ContentTemplate> 
<Triggers>

</Triggers> 
</asp:UpdatePanel> 
    







    
</asp:Content>


