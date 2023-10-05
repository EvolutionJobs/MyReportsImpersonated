Imports System.Data

Partial Class _AuthCheck
	Inherits System.Web.UI.Page

	Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load

		AuthMethod.Text = "Anonymous"
		AuthUser.Text = "none"
		Dim AUTH_TYPE As String = Request.ServerVariables("AUTH_TYPE")
		If AUTH_TYPE.Length > 0 Then AuthMethod.Text = AUTH_TYPE
		Dim AUTH_USER As String = Request.ServerVariables("AUTH_USER")
		If AUTH_USER.Length > 0 Then AuthUser.Text = AUTH_USER

		If AuthMethod.Text = "Negotiate" Then
			Dim auth As String = Request.ServerVariables.[Get]("HTTP_AUTHORIZATION")

			If (auth IsNot Nothing) AndAlso (auth.Length > 1000) Then
				AuthMethod.Text = AuthMethod.Text & " (KERBEROS)"
			Else
				AuthMethod.Text = AuthMethod.Text & " (NTLM)"
			End If
		End If

		ThreadId.Text = System.Security.Principal.WindowsIdentity.GetCurrent().Name
		dumpObject(System.Security.Principal.WindowsIdentity.GetCurrent(), tblProcessIdentity)
		dumpObject(System.Threading.Thread.CurrentPrincipal.Identity, tblThreadIdentity)
		Dim loop1, loop2 As Integer
		Dim coll As NameValueCollection
		coll = Request.ServerVariables
		Dim arr1 As String() = coll.AllKeys

		For loop1 = 0 To arr1.Length - 1
			Dim r As TableRow = New TableRow()
			Dim k As TableCell = New TableCell()
			Dim v As TableCell = New TableCell()
			k.Text = arr1(loop1)
			Dim arr2 As String() = coll.GetValues(arr1(loop1))

			For loop2 = 0 To arr2.Length - 1
				v.Text = v.Text + Server.HtmlEncode(arr2(loop2))
			Next

			v.Text = Server.HtmlEncode(arr2(0))
			r.Cells.AddRange(New TableCell() {k, v})
			tblSrvVar.Rows.Add(r)
		Next

		If Me.IsPostBack = False Then




		End If

	End Sub

	Protected Sub dumpObject(ByVal o As Object, ByVal outputTable As Table)
		Try
			Dim refl_WindowsIdenty As Type = GetType(System.Security.Principal.WindowsIdentity)
			Dim refl_WindowsIdenty_members As System.Reflection.MemberInfo() = o.[GetType]().FindMembers(System.Reflection.MemberTypes.[Property], System.Reflection.BindingFlags.[Public] Or System.Reflection.BindingFlags.Instance, Function(ByVal objMemberInfo As System.Reflection.MemberInfo, ByVal objSearch As Object) True, Nothing)

			For Each currentMemberInfo As System.Reflection.MemberInfo In refl_WindowsIdenty_members
				Dim r As TableRow = New TableRow()
				Dim k As TableCell = New TableCell()
				Dim v As TableCell = New TableCell()
				Dim getAccessorInfo As System.Reflection.MethodInfo = (CType(currentMemberInfo, System.Reflection.PropertyInfo)).GetGetMethod()
				k.Text = currentMemberInfo.Name
				Dim value As Object = getAccessorInfo.Invoke(o, Nothing)

				If GetType(IEnumerable).IsInstanceOfType(value) AndAlso Not GetType(String).IsInstanceOfType(value) Then

					For Each item As Object In CType(value, IEnumerable)
						v.Text += item.ToString() & "<br />"
					Next
				Else
					v.Text = value.ToString()
				End If

				r.Cells.AddRange(New TableCell() {k, v})
				outputTable.Rows.Add(r)
			Next

		Catch
		End Try
	End Sub

End Class
