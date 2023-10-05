<%@ Control Language="VB" AutoEventWireup="false" CodeFile="ucContactDetails.ascx.vb" Inherits="UserControls_ucContactDetails" %>





<center>
	<p><asp:Label ID="lblEditIntro" runat="server"></asp:Label></p>
	<p style="font-weight:bold;">Please make sure you enter your password before clicking the "Update & Verify" button</p>
</center>


<table width="100%">
	<tr><td align="center" colspan="2" style="text-decoration: underline;">Update Mobile Phone</td></tr>
	<tr><td colspan="2">Please select an International dialling code from the drop down based on where your mobile is registered and then enter your mobile number without the first zero.</td></tr>
	<tr><td align="right">Select Code:</td><td><asp:DropDownList ID="ddlstCodes" runat="server"></asp:DropDownList></td></tr>
	<tr><td align="right">Enter Phone:</td><td><asp:TextBox ID="txtEditPhone" runat="server" Width="150px"></asp:TextBox></td></tr>
	<tr><td align="right" style="color:darkgreen;">Enter Password:</td><td><asp:TextBox ID="txtPasswordPhone" runat="server"  TextMode="Password" ></asp:TextBox></td></tr>
	<tr><td></td><td><asp:Button ID="cmdUpdatePhone" runat="server" Text="Update & Verify" /></td></tr>
	<tr><td>&nbsp;</td><td>&nbsp;</td></tr>
	<tr><td  align="center"colspan="2" style="text-decoration: underline;">Update Personal Email Address</td></tr>
	<tr><td align="right" colspan="2" >Please enter a "Personal Email Address". For security reasons, this cannot be a work email address.</td></tr>
	<tr><td align="right">Enter Email:</td><td><asp:TextBox ID="txtEditEmail" runat="server" Width="300px" ></asp:TextBox></td></tr>
	<tr><td align="right" style="color:darkgreen;">Enter Password:</td><td><asp:TextBox ID="txtPasswordEmail" runat="server"   TextMode="Password" ></asp:TextBox></td></tr>
	<tr><td></td><td><asp:Button ID="cmdUpdateEmail" runat="server" Text="Update & Verify" /></td></tr>
</table>

<br />
<center>
<asp:Label ID="lblErrors" runat="server" ForeColor="Red" Font-Bold="true"></asp:Label>
</center>
<br />
<br />
<br />
<br />
<br />
<br />
<br />
<br />
<br />
<br />
<br />
<br />
<br />
<center>
<asp:LinkButton ID="lnkSMSTest" runat="server" Visible="False" >...</asp:LinkButton>
<br />
<asp:Label ID="lblSMSTestError" runat="server" ForeColor="Red" Font-Bold="true"></asp:Label>
<br />
</center>
<br />



