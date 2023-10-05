<%@ Page Title="" Language="VB" MasterPageFile="~/Site.master" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="_Default" %>

<asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" Runat="Server">
<center>
<asp:Label ID="lblWindowsUsername" runat="server" Text=""></asp:Label>
<asp:Panel ID="pnlLogin" runat="server">

    Enter password to access Wages pages:
    <br />
    <asp:TextBox ID="txtPassword" runat="server" TextMode="Password"></asp:TextBox>
    <br />
    <asp:Button ID="cmdSubmit" runat="server" Text="Login" />
    <br />
    <asp:LinkButton ID="lnkcmdChangePassword" runat="server">Change Password</asp:LinkButton>

</asp:Panel>
<asp:Panel ID="pnlChangePassword" runat="server">
<span style='text-decoration: underline'>Change Password</span>
    <br />
<table>
<tr><td align="right">Old Password:</td><td align="left"><asp:TextBox ID="txtOldPassword" runat="server" TextMode="Password"></asp:TextBox></td></tr>
<tr><td align="right">New Password:</td><td align="left"><asp:TextBox ID="txtNewPassword1" runat="server" TextMode="Password"></asp:TextBox></td></tr>
<tr><td align="right">Retype New Password:</td><td align="left"><asp:TextBox ID="txtNewPassword2" runat="server" TextMode="Password"></asp:TextBox></td></tr>
<tr><td align="right"></td><td align="left"><asp:Button ID="cmdChangePassword" runat="server" Text="Login" /></td></tr>
</table>


    
    <br />
    

</asp:Panel>
    <br />
    <asp:Label ID="lblStatus" runat="server" Text="" ForeColor="Red" Font-Bold="true"></asp:Label>

        <br /><br /><br />
</center>



<div style="display:none">
        <br /><br /><br />
            <br />
        <asp:Label ID="lblURL" runat="server" Text=""></asp:Label>
    <br />
    <asp:TextBox ID="txtTest" runat="server"></asp:TextBox>
    <br />
    <asp:Button ID="cmdTest" runat="server" Text="TestSwap" />
        <br />
        <asp:Label ID="lblTest" runat="server" Text=""></asp:Label>

</div>



</asp:Content>

