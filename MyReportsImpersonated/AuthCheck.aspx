<%@ Page Title="" Language="VB"  AutoEventWireup="false" CodeFile="AuthCheck.aspx.vb" Inherits="_AuthCheck" %>
<script runat="server">





</script>

<html>
<body>
<font face="Verdana">
<center><H2>WHO Page</H2><br>
<table border=1>
<tr>
<td><img src="img1.gif"></td>
<td><img src="img2.gif"></td>
<td><img src="img3.gif"></td>
<td><img src="img4.gif"></td>
<td><img src="img5.gif"></td>
<td><img src="img6.gif"></td>
<td><img src="img7.gif"></td>
<td><img src="img8.gif"></td>
<td><img src="img9.gif"></td>
<td><img src="img10.gif"></td>
</tr>
</table>
</center>
<table border=1>
<tr><td>Authentication Method </td><td><b><asp:label id=AuthMethod runat=server/></b></td><td>Request.ServerVariables("AUTH_TYPE")</td></tr>
<tr><td>Identity </td><td><b><asp:label id=AuthUser runat=server /></b></td><td>Request.ServerVariables("AUTH_USER") or System.Threading.Thread.CurrentPrincipal.Identity</td></tr>
<tr><td>Windows identity </td><td> <b><asp:label id=ThreadId runat=server /></b></td><td>System.Security.Principal.WindowsIdentity.Getcurrent</td></tr>
</table>

<fieldset>
            <label>Identity (System.Threading.Thread.CurrentPrincipal.Identity)</label>
            <asp:Table ID="tblThreadIdentity" runat="server"></asp:Table>
        </fieldset>


	<fieldset>
            <label>Windows Identity (System.Security.Principal.WindowsIdentity.GetCurrent)</label>
            <asp:Table ID="tblProcessIdentity" runat="server"></asp:Table>
        </fieldset>

Dump of server variables :
<br><br>
<asp:Table ID="tblSrvVar" runat="server"></asp:Table>

<form name="form1" method="POST" action="default.aspx">
Sample Form<br>
<input NAME="sometext" TYPE="Text" SIZE="24">
<input TYPE="submit" VALUE="Submit" NAME="SubmitForm" > 
</form>

</font>
</body>
</html> 

