
Partial Class _Error
	Inherits System.Web.UI.Page
	Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load


		If Request("e") <> "" Then


			lblError.Text = "" & Request("e")

			Dim sReferer As String = ""
			If Not HttpContext.Current.Request.ServerVariables("HTTP_REFERER") = Nothing Then
				sReferer = "HTTP_REFERER=" & HttpContext.Current.Request.ServerVariables("HTTP_REFERER")
			End If

			SaveToLog("ERROR__PAGE___ [lblError.Text=" & lblError.Text & "]; Request.Url.ToString=" & Request.Url.ToString & "; sReferer=" & sReferer)


		End If


	End Sub

End Class
