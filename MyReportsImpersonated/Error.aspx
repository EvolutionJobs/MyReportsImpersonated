<%@ Page Title="" Language="VB" MasterPageFile="~/Site.master" AutoEventWireup="false" CodeFile="Error.aspx.vb" Inherits="_Error" %>

<asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" Runat="Server">
	<style type="text/css">
		.auto-style1 {
			height: 19px;
		}
		.auto-style2 {
			color: red;
		}
	</style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" Runat="Server">

<center>

        <div>

			<u><p><h1>Error:</h1></u>
			<p>
				<asp:Label ID="lblError" runat="server" ForeColor="Red"></asp:Label>
			</p>

<table>
<tr><td align="right" class="auto-style1">(1) </td><td align="left" class="auto-style1">Please click this link: <a href="http://EvoWeb3/LoginWindows.aspx">http://EvoWeb3/LoginWindows.aspx</a> to try and login again using your <span class="auto-style2">Windows credentials</span></td></tr>
<tr><td align="right" class="auto-style1">&nbsp;</td><td align="left" class="auto-style1"></td></tr>
<tr><td align="right">(2) </td><td align="left">If that doesn&#39;t work, then please try accessing My Reports from a browser you don&#39;t normally use, e.g. Edge:</td></tr>
<tr><td align="right"></td><td align="left">Please copy <a href="http://evoweb3">http://evoweb3</a> and paste into Edge and login
				and see if that works before contacting support</td></tr>
<tr><td align="center" colspan="2">&nbsp;</td></tr>
<tr><td align="right">(3) </td><td align="left">And if that doesn&#39;t work, then please try clearing the browser cache:</td></tr>
<tr><td align="right" class="auto-style1">&nbsp;</td><td align="left" class="auto-style1">(a) If you are using Chrome, then open Chrome</td></tr>
<tr><td align="right" class="auto-style1">&nbsp;</td><td align="left" class="auto-style1">(b) At the top right, click More (the 3 dots)</td></tr>
<tr><td align="right" class="auto-style1">&nbsp;</td><td align="left" class="auto-style1">(c) Click &#39;More tools&#39; &gt; &#39;Clear browsing data&#39;</td></tr>
<tr><td align="right" class="auto-style1">&nbsp;</td><td align="left" class="auto-style1">or Click &#39;Settings&#39; &gt; &#39;Privacy and Security&#39; &gt;&nbsp; &#39;Clear browsing data&#39;</td></tr>
<tr><td align="right" class="auto-style1">&nbsp;</td><td align="left" class="auto-style1">(d) At the top, choose a time range. To delete everything, select &#39;All time&#39;</td></tr>
<tr><td align="right" class="auto-style1">&nbsp;</td><td align="left" class="auto-style1">(e) Next to &#39;Cookies and other site data&#39;and &#39;Cached images and files&#39;, check the boxes</td></tr>
<tr><td align="right" class="auto-style1">&nbsp;</td><td align="left" class="auto-style1">(f) Click Clear data</td></tr>
<tr><td align="right" class="auto-style1">&nbsp;</td><td align="left" class="auto-style1">(g) Then try logging on again</td></tr>
<tr><td align="center" colspan="2">&nbsp;</td></tr>
<tr><td align="right">(4)</td><td align="left">If that still does not work, then please contact support</td></tr>
<tr><td align="center" colspan="2"></td></tr>

</table>

			<p>
				<br /><br />
				<br /></p>
			<p></p>
	
        </div>

</center>


</asp:Content>


