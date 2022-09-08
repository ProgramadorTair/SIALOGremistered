
Public Class Form1
    Public Conex As New ADODB.Connection
    Public CONEXCRYSREPOR As String
    Dim lOrigen As String
    Dim lOrigenSpeLib As String
    Dim lrsX As New ADODB.Recordset
    Private Sub Guna2GradientButton1_Click(sender As Object, e As EventArgs) Handles Guna2GradientButton1.Click
        Conex.CommandTimeout = 50
        lOrigen = "Provider=MSDASQL.1;Persist Security Info=False;User ID=" + Guna2TextBox2.Text + ";PWD=" + Guna2TextBox3.Text + ";Data Source=SPEED400" + Guna2TextBox1.Text + ";SPEEDPTF6"
        Conex.Open(lOrigen)
        If Err.Number <> 0 Then
            MsgBox(Err.Number & Chr(13) & Err.Description, vbExclamation, "Mensaje del Sistema")
        Else
            'lrsX.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            'lrsX.CursorType = ADODB.CursorTypeEnum.adOpenStatic
            'lrsX.LockType = ADODB.LockTypeEnum.adLockReadOnly
            'lrsX.ActiveConnection = Conex
            'lOrigenSpeLib = "select * from talma"
            'lrsX.Open(lOrigenSpeLib)
            'MsgBox(lrsX.RecordCount)
            MsgBox("Se conecto correctamente")
            System.Windows.Forms.Form.ActiveForm.Hide()
            Menulog.Show()
        End If
    End Sub
End Class
