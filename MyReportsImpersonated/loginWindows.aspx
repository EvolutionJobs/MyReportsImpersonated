<%@ Page Title="" Language="VB" MasterPageFile="~/Site.master" AutoEventWireup="false" CodeFile="loginWindows.aspx.vb" Inherits="WagesLoginWindows" %>

<asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" Runat="Server">

<center>


<asp:Panel ID="pnlLogin" runat="server">

	<center><h2>Windows Login Password</h2></center>

	<p>Enter your WINDOWS password that you use to login to your PC/Laptop with:</p>
    
    <p>
	<asp:Table runat="server">
		<asp:TableRow>
			<asp:TableCell>
				Windows Username=
			</asp:TableCell>
			<asp:TableCell>
				<asp:TextBox ID="txtWindowsUsername" runat="server"  ReadOnly="true" Width="200px" BackColor="lightyellow"></asp:TextBox>
			</asp:TableCell>
		</asp:TableRow>
		<asp:TableRow>
			<asp:TableCell>
				Windows Password=
			</asp:TableCell>
			<asp:TableCell>
				<asp:TextBox ID="txtWindowsPassword" runat="server" TextMode="Password" Width="200px" BackColor="lightyellow"></asp:TextBox>
			</asp:TableCell>
		</asp:TableRow>
		<asp:TableRow>
			<asp:TableCell>
				
			</asp:TableCell>
			<asp:TableCell HorizontalAlign="Left">
				<asp:Button ID="cmdSubmit" runat="server" Text="Login with Windows Credentials" Width="200px" />
			</asp:TableCell>
		</asp:TableRow>

	</asp:Table>
	
    
    </p>
    
    <p>
    
    </p>
	<p><asp:Label ID="lblStatus" runat="server" Text="" ForeColor="Red" Font-Bold="true"></asp:Label></p>

  
</asp:Panel>

        <br /><br /><br />
</center>

	   

</asp:Content>

