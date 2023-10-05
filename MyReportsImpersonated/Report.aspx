<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Report.aspx.vb" Inherits="Report" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=15.0.0.0, Culture=neutral, PublicKeyToken=89845dcd8080cc91" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="X-UA-Compatible" content="IE=edge" >
    <title></title>
<link href="~/Styles/Report.css" rel="stylesheet" type="text/css" />
<style>
h1   {
	height: 52px;
	display: flex;
	align-items: center;
	padding: 0 12px;
	background-color: #3f51b5;
	color: #fff;
	font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Helvetica Neue', Arial, sans-serif;
	font-weight: 300;
	margin: 0;
	font-size: 30px;
	overflow: hidden;
	box-shadow: rgba(0, 0, 0, 0.14) 0px 0px 4px, rgba(0, 0, 0, 0.28) 0px 4px 8px;
	z-index: 1;
}
</style>
<script src="js/jquery-3.3.1.min.js"></script>
<script type="text/javascript">

$(document).ready(function () {
    //Disable cut copy paste
    $('body').bind('cut copy paste', function (e) {
        e.preventDefault();
    });
   
    //Disable mouse right click
    $("body").on("contextmenu",function(e){
        return false;
    });
});

</script>
</head>
<body oncontextmenu="return false" onselectstart="return false"
      onkeydown="if ((arguments[0] || window.event).ctrlKey) return false">
	<h1>My Reports</h1>
    <form id="form1" runat="server">
    <asp:scriptmanager ID="Scriptmanager2" runat="server"></asp:scriptmanager>
    <div>

        <rsweb:ReportViewer ID="ReportViewer1" runat="server" Font-Names="Verdana" 
            Font-Size="8pt" InteractiveDeviceInfos="(Collection)" ProcessingMode="Remote" 
            WaitMessageFont-Names="Verdana" WaitMessageFont-Size="14pt" Width="100%" Height="800" 
            ShowPrintButton="false" ShowExportControls="false" 
            ShowBackButton="true" ShowToolBar="true"  CssClass="cssReport"  >
            <ServerReport ReportPath=""  ReportServerUrl=""  />
        </rsweb:ReportViewer>

    </div>
    </form>

</body>
</html>
