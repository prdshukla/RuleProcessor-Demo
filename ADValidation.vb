Imports System.DirectoryServices.AccountManagement

Public Class ADValidation

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim Success As Boolean = False
        Dim Entry As New System.DirectoryServices.DirectoryEntry("LDAP://" & Trim(TextBox1.Text), Trim(TextBox2.Text), Trim(TextBox3.Text))
        Dim Searcher As New System.DirectoryServices.DirectorySearcher(Entry)

        Searcher.SearchScope = DirectoryServices.SearchScope.OneLevel

        Try
            Dim Results As System.DirectoryServices.SearchResult = Searcher.FindOne
            Success = Not (Results Is Nothing)
        Catch
            Success = False
        End Try

        MsgBox(Success)

    End Sub
End Class