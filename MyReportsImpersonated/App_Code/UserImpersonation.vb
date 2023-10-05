Imports System.Security.Principal
Imports System.Runtime.InteropServices

Public Class UserImpersonation
	<DllImport("advapi32.dll")>
	Public Shared Function LogonUserA(ByVal lpszUserName As [String], ByVal lpszDomain As [String], ByVal lpszPassword As [String], ByVal dwLogonType As Integer, ByVal dwLogonProvider As Integer, ByRef phToken As IntPtr) As Integer
	End Function

	<DllImport("advapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
	Public Shared Function DuplicateToken(ByVal hToken As IntPtr, ByVal impersonationLevel As Integer, ByRef hNewToken As IntPtr) As Integer
	End Function

	<DllImport("advapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
	Public Shared Function RevertToSelf() As Boolean
	End Function

	<DllImport("kernel32.dll", CharSet:=CharSet.Auto)>
	Public Shared Function CloseHandle(ByVal handle As IntPtr) As Boolean
	End Function

	Public Const LOGON32_LOGON_INTERACTIVE As Integer = 2
	Public Const LOGON32_PROVIDER_DEFAULT As Integer = 0

	Private impersonationContext As WindowsImpersonationContext

	'Private Const UserName As String = "USER_ID"
	'Private Const Password As String = "USER_DOMAIN_PASSWORD"
	'Private Const Domain As String = "USER_DOMAIN_NAME"

	Public Function ImpersonateValidUser(ByVal UserName As String, ByVal EncryptedPassword As String, ByVal Domain As String) As Boolean

		Dim sDecryptedPass As String = clsDTEmailAccount.clsAccountDetails.CryptoDecrypt(EncryptedPassword, "support@evolutionjobs.co.uk")


		Dim tempWindowsIdentity As WindowsIdentity
		Dim token As IntPtr = IntPtr.Zero
		Dim tokenDuplicate As IntPtr = IntPtr.Zero
		If RevertToSelf() Then
			If LogonUserA(UserName, Domain, sDecryptedPass, LOGON32_LOGON_INTERACTIVE, LOGON32_PROVIDER_DEFAULT, token) <> 0 Then
				If DuplicateToken(token, 2, tokenDuplicate) <> 0 Then
					tempWindowsIdentity = New WindowsIdentity(tokenDuplicate)
					impersonationContext = tempWindowsIdentity.Impersonate()
					If impersonationContext IsNot Nothing Then
						CloseHandle(token)
						CloseHandle(tokenDuplicate)
						Return True
					End If
				End If
			End If
		End If
		If token <> IntPtr.Zero Then
			CloseHandle(token)
		End If
		If tokenDuplicate <> IntPtr.Zero Then
			CloseHandle(tokenDuplicate)
		End If
		Return False
	End Function

	Public Sub UndoImpersonation()
		'If impersonationContext IsNot Nothing Then
		'	impersonationContext.Undo()
		'End If
	End Sub

End Class