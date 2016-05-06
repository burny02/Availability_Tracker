Imports TemplateDB

Module Variables
    Public OverClass As OverClass
    Private Const TablePath As String = "\\hvivo.sharepoint.com@SSL\DavWWWRoot\hSite\Systems\Availability.accdb"
    Private Const PWord As String = "RetroRetro*1"
    Private Const Connect2 As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & TablePath & ";Jet OLEDB:Database Password=" & PWord
    Private Const UserTable As String = "[Users]"
    Private Const UserField As String = "Username"
    Private Const AuditTable As String = "[Audit]"
    Private Contact As String = "Daniel Ramsay"
    Public Const SolutionName As String = "Availability Tracker"
    Public Bank As Boolean = False
    Public StaffID As Long = 0

    Public Function GetTheConnection() As String
        GetTheConnection = Connect2
    End Function


    Public Sub StartUp()

        OverClass = New TemplateDB.OverClass
        OverClass.SetPrivate(UserTable,
                                   UserField,
                                   Nothing,
                                   Contact,
                                   Connect2,
                                   AuditTable)

        OverClass.LoginCheck()

        OverClass.AddAllDataItem(Form1)

        Dim dt As DataTable = OverClass.TempDataTable("SELECT Bank, StaffID FROM Users WHERE UserName='" & OverClass.GetUserName & "'")

        If dt.Rows.Count = 0 Then
            MsgBox("An error occurred.")
            Application.Exit()
        End If

        Bank = dt.Rows(0).Item(0)
        StaffID = dt.Rows(0).Item(1).ToString

        If StaffID = 0 Then
            MsgBox("An error occurred.")
            Application.Exit()
        End If

    End Sub

End Module
