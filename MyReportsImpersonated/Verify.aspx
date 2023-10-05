<%@ Page Title="" Language="VB" MasterPageFile="~/Site.master" AutoEventWireup="false" CodeFile="Verify.aspx.vb" Inherits="_Verify" %>

<asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" Runat="Server">
    </asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" Runat="Server">







    <table width="100%">
    <tr>
        <td align="center">
            <asp:Label ID="lblEditIntro" runat="server"></asp:Label>
        </td>
    </tr>

    <tr>
        <td align="center">
        <asp:TextBox ID="txtCode" runat="server" Width="277px"></asp:TextBox>
        </td>
    </tr>
    <tr>
        <td align="center">
        <asp:Button ID="cmdVerify" runat="server" Text="Submit" Width="67px" />
        </td>
        </tr>
    <tr>
        <td align="center"><asp:Label ID="lblStatus" runat="server" Text="" ForeColor="Red" Font-Bold="true"></asp:Label></td>
        </tr>
    <tr>
        <td align="center">
         <asp:LinkButton ID="lnkcmdLoginPage" runat="server">Login Page</asp:LinkButton>&nbsp;|&nbsp;<asp:LinkButton ID="lnkcmdManageMyAccount" runat="server">Manage My Account</asp:LinkButton>
        </td>
    </tr>

    </table>    


    <br />
    

        <br /><br /><br />




<div style="display:none">
<asp:Label ID="lblWindowsUsername" runat="server" Text="" ></asp:Label>
    <asp:Label ID="lblType" runat="server" Text="" ></asp:Label>
</div>



</asp:Content>

