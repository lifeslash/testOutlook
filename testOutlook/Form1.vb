Imports Outlook = Microsoft.Office.Interop.Outlook

Public Class Form1

    Private oApp As Outlook.Application 'outlook application
    Private oNS As Outlook.NameSpace    'outlook namespace, MAPI means parameter when use Namespaces
    Private oContactsFolder As Outlook.MAPIFolder   '

    Private Sub form_laod(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        oApp = New Outlook.Application()
        oNS = oApp.GetNamespace("MAPI")

    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button1.Click
        oContactsFolder = oNS.PickFolder()  'the return is object of MAPIFolder type

    End Sub

End Class
