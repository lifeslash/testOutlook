Imports Outlook = Microsoft.Office.Interop.Outlook

Public Class Form1

    Public oApp As Outlook.Application
    Public oNS As Outlook.NameSpace


    Private Sub form_laod(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        oApp = New Outlook.Application()
        oNS = oApp.GetNamespace("MAPI")

    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button1.Click

    End Sub

End Class
