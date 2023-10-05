<%@ Page Title="" Language="VB" MasterPageFile="~/Site.master" AutoEventWireup="false" CodeFile="login.aspx.vb" Inherits="WagesLogin" %>

<asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" Runat="Server">

<center>
<asp:Label ID="lblWindowsUsername" runat="server" Text=""></asp:Label>

<asp:Panel ID="pnlDefault" runat="server" Visible =" false">

	<center><h2>No Report Selected</h2></center>

	<p>Please select a report from the "My Reports" page</p>

</asp:Panel>


<asp:Panel ID="pnlLogin" runat="server">

	<center><h2>Login</h2></center>

	<p>Enter your <span style="color:red;">Wages Pages password</span> to access Reports pages:<br />

		<table style="">
		<tr><td align="center" colspan="2">IMPORTANT</td></tr>
		<tr><td align="left">(a)</td><td align="left">This should <u>not</u> be your Windows password. It should be a personal password that only you know</td></tr>
		<tr><td align="left">(b)</td><td align="left">If you don't have a password, then please contact <a href='mailto:data.team@evolutionjobs.co.uk'>data.team@evolutionjobs.co.uk</a></td></tr>
		<tr><td align="left">(c)</td><td align="left">If you have forgotten your password, then please click the <b>Forgotten Password</b> link below</td></tr>
		<tr><td align="left"></td><td align="left"></td></tr>
		</table>

	</p>

    <p>
    <asp:TextBox ID="txtPassword" runat="server" TextMode="Password"></asp:TextBox>
    </p>
    
    <p>
    <asp:Button ID="cmdSubmit" runat="server" Text="Login" />
    </p>


    <p>
		<asp:LinkButton ID="lnkcmdChangePassword" runat="server" >Change Password</asp:LinkButton>
		<asp:Literal ID="litPipe1" runat="server" Text="&nbsp;|&nbsp;"></asp:Literal>
		<asp:LinkButton ID="lnkcmdFogottenPassword" runat="server">Forgotten Password</asp:LinkButton>
		<asp:Literal ID="litPipe2" runat="server" Text="&nbsp;|&nbsp;"></asp:Literal>
		<asp:LinkButton ID="lnkcmdRequestNewPassword" runat="server">Request New Password</asp:LinkButton> <br />
        
    </p>
   

</asp:Panel>


<asp:Panel ID="pnlChangePassword" runat="server">
	<center><h2>Change Password</h2></center>

	<p><asp:Label ID="lblSpecialMessage" runat="server" Font-Bold="true"></asp:Label></p>

<table>
<tr><td align="right"><asp:Label ID="lblOldPassword" runat="server"></asp:Label></td><td align="left"><asp:TextBox ID="txtOldPassword" runat="server" TextMode="Password"></asp:TextBox></td></tr>
<tr><td align="right">New Password:</td><td align="left"><asp:TextBox ID="txtNewPassword1" runat="server" TextMode="Password"></asp:TextBox></td></tr>
<tr><td align="right">Retype New Password:</td><td align="left"><asp:TextBox ID="txtNewPassword2" runat="server" TextMode="Password"></asp:TextBox></td></tr>
<tr><td align="right"></td><td align="left"><asp:Button ID="cmdChangePassword" runat="server" Text="Change Password" /></td></tr>
<tr><td align="center" colspan="2"></td></tr>
</table>

	<p>

<asp:LinkButton ID="lnkcmdMyReports" runat="server">My Reports</asp:LinkButton>&nbsp;|&nbsp;<asp:LinkButton ID="lnkcmdLogin" runat="server">Login</asp:LinkButton>
	</p>

</asp:Panel>


		<asp:LinkButton ID="lnkcmdManageMyAccount" runat="server">Manage My Account</asp:LinkButton>
    <p><asp:Label ID="lblStatus" runat="server" Text="" ForeColor="Red" Font-Bold="true"></asp:Label></p>
		<asp:Label ID="lblURL" runat="server"></asp:Label>
        <br /><br /><br />
</center>



<div style="display:none">
        <br /><br /><br />

   
            <br />
       <asp:LinkButton ID="lnkbackToLoginPage" runat="server">Login Page</asp:LinkButton>
    <br />
    <asp:TextBox ID="txtTest" runat="server"></asp:TextBox>
    <br />
    <asp:Button ID="cmdTest" runat="server" Text="TestSwap" />
        <br />
        <asp:Label ID="lblTest" runat="server" Text=""></asp:Label>
        <br />
        <asp:Label ID="lblTemp" runat="server" Text=""></asp:Label>
	<asp:Label ID="lblCurrentURL" runat="server" Text=""></asp:Label>
</div>


</asp:Content>

