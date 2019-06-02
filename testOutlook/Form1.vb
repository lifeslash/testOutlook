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

        Dim filter As String = "[MessageClass] = ""IPM.Contact"""
        Dim oContactItems As Outlook.Items
        oContactItems = oContactsFolder.Items.Restrict(filter)

        For Each oContact As Outlook.ContactItem In oContactItems

            'ContactItemクラス
            'LastName、FirstName、MiddleName、Title
            'JobTitle
            'Email1Address、Email1DisplayName、IMAddress
            'WebPage
            'BusinessTelephoneNumber、OtherTelephoneNumber、PagerNumber、MobileTelephoneNumber、BusinessFaxNumber

            Dim item As ListViewItem
            item.Text = oContact.EmailDisplayName
            item.Tag = oContact
            ListBox1.Items.Add(item)
        Next
    End Sub

End Class
