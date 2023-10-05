<%@ Page Title="" Language="VB" MasterPageFile="~/Site.master" AutoEventWireup="false" CodeFile="ForgottenPassword.aspx.vb" Inherits="_ForgottenPassword" %>

<%@ Register src="UserControls/ucContactDetails.ascx" tagname="ucContactDetails" tagprefix="uc1" %>

<asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" Runat="Server">

<center><h2><asp:Label ID="lblHeading" runat="server" Text=""></asp:Label></h2></center>

<asp:Panel ID="pnlNewPasswordStep1" runat="server">
	<center>
	<p>
	If this is the first time you are accessing the site and require a New Password, then please <b>CALL HR</b> and provide them with
	either a mobile phone number and/or personal email address. 
	</p>
	<p>
		After they have entered the details into the system, click the Refreh button below.This will show the details they entered. 
		To keep your details secure, you will need to click the "Verify" button to verify the contact information before proceeding.
	</p>
	<asp:LinkButton ID="lnkcmdRefreshFromStep1" runat="server">Refresh</asp:LinkButton>
	</center>
</asp:Panel>

<asp:Panel ID="pnlNewPasswordStep2" runat="server">
	<p>
		<asp:Literal ID="litContactIntro" runat="server"></asp:Literal>
	</p>


    <table width="100%" border="1" style=" border-collapse: collapse; border: 1px solid gray">
    <tr><td align="left" valign="top" width="15%"></td>
		<td align="left" valign="top" width="38%">
        	&nbsp;</td>
        <td align="center" valign="top" width="25%">Verification</td>
		<td align="center" valign="top" width="22%">Temporary Password</td>
    </tr>
    <tr><td align="left"  valign="top">Mobile:</td>
		<td align="left" valign="top">
        <asp:Label ID="lblMobile" runat="server" Text=""></asp:Label>
        </td>
        <td align="center"  valign="top">
            <asp:LinkButton ID="lnkcmdVerifyMobile" runat="server">Verify</asp:LinkButton>&nbsp;&nbsp;<asp:LinkButton ID="lnkcmdReSendVerifyCodeMobile" runat="server">ReSend Code</asp:LinkButton>
        </td>
		<td align="center"  valign="top"><asp:LinkButton ID="lnkcmdSendPasswordToMobile" runat="server">Send Password</asp:LinkButton></td>
    </tr>
    <tr><td align="left" valign="top">Personal Email:</td>
		<td align="left" valign="top">
        <asp:Label ID="lblEmail" runat="server" Text=""></asp:Label>
        </td>
        <td align="center" valign="top">
            <asp:LinkButton ID="lnkcmdVerifyEmail" runat="server">Verify</asp:LinkButton>&nbsp;&nbsp;<asp:LinkButton ID="lnkcmdReSendVerifyCodeEmail" runat="server">ReSend Code</asp:LinkButton>			
        </td>
		<td align="center" valign="top"><asp:LinkButton ID="lnkcmdSendPasswordToEmail" runat="server">Send Password</asp:LinkButton></td>
        </tr>

    </table>
	<br />
	<center>
		<asp:Button ID="cmdEditDetails" runat="server" Text="Edit Details" Width="131px" /><br />
		<asp:LinkButton ID="lnkcmdRefreshFromStep2" runat="server">Refresh</asp:LinkButton>
	</center>

</asp:Panel>

<uc1:ucContactDetails ID="ucContactDetails1" runat="server" Visible="false" />

	<center><asp:LinkButton ID="lnkcmdCancelEdit" runat="server">Cancel</asp:LinkButton></center>


<asp:Panel ID="pnlNewPasswordStep3" runat="server">
	<center>
    <p>
        Please enter the code that was sent to your mobile or email:
    </p>
    <p>
        <asp:TextBox ID="txtVerificationCode" runat="server" Width="190px" ></asp:TextBox><br /><br />
        <asp:Button ID="cmdVerifyCode" runat="server" Text="Submit" /><br />
    </p>
	<asp:LinkButton ID="lnkcmdCancelVerify" runat="server">Cancel</asp:LinkButton>
	</center>
</asp:Panel>


	
    <br />
	<center>
		<asp:LinkButton ID="lnkcmdMyReports" runat="server">My Reports</asp:LinkButton>&nbsp;|&nbsp;<asp:LinkButton ID="lnkcmdLogin" runat="server">Login</asp:LinkButton><br /><br />
    <asp:Label ID="lblStatus" runat="server" Text="" ForeColor="Red" Font-Bold="true"></asp:Label><br />
	
	</center>

        <br /><br /><br />




<div style="display:none">
<asp:Label ID="lblWindowsUsername" runat="server" Text="" ></asp:Label>
	<asp:Label ID="lblTemp" runat="server" Text="" ></asp:Label>
</div>



</asp:Content>

