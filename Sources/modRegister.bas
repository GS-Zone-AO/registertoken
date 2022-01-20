Attribute VB_Name = "modRegister"
Option Explicit

Public Declare Function IsUserAnAdmin Lib "Shell32" () As Integer

Const ClientName As String = "Cliente de GS-Zone AO"
Const UrlCode As String = "gszoneao"

Private Sub RegisterClient(ByVal presentClient As String)

    CreateObject("WScript.Shell").RegWrite "HKCR\" & UrlCode & "\", ClientName
    CreateObject("WScript.Shell").RegWrite "HKCR\" & UrlCode & "\URL Protocol", "", "REG_SZ"
    CreateObject("WScript.Shell").RegWrite "HKCR\" & UrlCode & "\DefaultIcon\", """" & presentClient & """,-1"
    CreateObject("WScript.Shell").RegWrite "HKCR\" & UrlCode & "\shell\open\command\", """" & presentClient & """  ""%1"""
    
End Sub

Function FileExist(ByVal file As String, _
                   Optional FileType As VbFileAttribute = vbNormal) As Boolean
    '*****************************************************************
    'Se fija si existe el archivo
    '*****************************************************************

    FileExist = LenB(Dir$(file, FileType)) <> 0

End Function

Sub Main()

    Dim presentClient As String
    presentClient = App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "Argentum" & ".exe"
    If Not FileExist(presentClient) Then
        presentClient = App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "ClientGSZAO" & ".exe"
        If Not FileExist(presentClient) Then
            MsgBox "No se ha encontrado ningun ejecutable compatible con el cliente.", vbCritical + vbOKOnly
            End
        End If
    End If
    
    If IsUserAnAdmin() = 0 Then
        MsgBox "Se requieren permisos de Administrador para registrar el cliente.", vbCritical + vbOKOnly
    Else
        Call RegisterClient(presentClient)
        MsgBox "Cliente registrado con exito!", vbInformation + vbOKOnly
    End If

End Sub
