Public Class Form1

 Dim Thread1, Thread2, Thread3 As System.Threading.Thread
   Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
     Me.CheckForIllegalCrossThreadCalls = False
   End Sub
   Private Sub CheckProgress()
     BackgroundWorker1.RunWorkerAsync()
   End Sub
   Private Sub Search()
     Dim MainFolder = BrowsePathText.Text
     If MainFolder = "" Then
       MsgBox("Please Choose Directory")
       flpath = " "
       ListBox1.Items.Clear()
   Else
       Dim FolderInfo As New System.IO.DirectoryInfo(MainFolder)
       Files = FolderInfo.GetFiles("*" & SearchText.Text & "*.*",
    IO.SearchOption.AllDirectories)
      For Each Results In Files
        flpath = " "
       flpath = Results.Name
       ListBox1.Items.Add(flpath)
     Next
     Timer1.Enabled = False
     If ListBox1.Items.Count() = 0 Then
       MsgBox("No Result Found")
     End If
     End If
 End Sub

 Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
     Timer1.Enabled = True
     CheckProgress()
     Thread1 = New System.Threading.Thread(AddressOf Search)
     Thread1.Start()
     Timer1.Enabled = False
 End Sub

 Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
   Timer1.Enabled = True
   Dim NewFolder As System.IO.DirectoryInfo = New System.IO.DirectoryInfo(BrowsePathText.Text + "\" + SearchText.Text)
   Dim Items = ListBox1.Items.Count()
   If NewFolder.Exists Then
      MsgBox("Folder Already Created")
      For Each number As String In ListBox1.Items
         Dim check As System.IO.DirectoryInfo = New System.IO.DirectoryInfo(BrowsePathText.Text + "\" + SearchText.Text + "\" + ListBox1.GetItemText(number))
      If check.Exists Then
         MsgBox("Files Already In It")
      Else
         Try
             System.IO.File.Move(BrowsePathText.Text + "\" + ListBox1.GetItemText(number), BrowsePathText.Text + "\" + SearchText.Text + "\" + ListBox1.GetItemText(number))
         Catch damn As System.IO.FileNotFoundException
            Form2.Show()
            Form2.Close()
        End Try
     End If
   Next
   MsgBox("Folder Created " & "Files are moved", MsgBoxStyle.Information)
Else
 NewFolder.Create()
 For Each number As String In ListBox1.Items
   Dim check As System.IO.DirectoryInfo = New System.IO.DirectoryInfo(BrowsePathText.Text + "\" + SearchText.Text + "\" +
ListBox1.GetItemText(number))
 If check.Exists Then
     MsgBox("Files Already In It")
 Else
   Try
 System.IO.File.Move(BrowsePathText.Text + "\" + ListBox1.GetItemText(number), BrowsePathText.Text + "\" + SearchText.Text + "\" + ListBox1.GetItemText(number))
 Catch damn As System.IO.FileNotFoundException
   Form2.Show()
   Form2.Close()
 End Try
 End If
 Next
 MsgBox("Folder Created " & "Files are moved", MsgBoxStyle.Information)
 End If
 End Sub
 Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs)
Handles ListBox1.SelectedIndexChanged
 End Sub
 Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles
SearchText.TextChanged
 End Sub
 Private Sub FolderBrowserDialog1_HelpRequest(sender As Object, e As EventArgs)
Handles FolderBrowserDialog1.HelpRequest
 End Sub
 Private Sub Button2_Click(sender As Object, e As EventArgs) Handles
Button2.Click
 If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
 BrowsePathText.Text = FolderBrowserDialog1.SelectedPath
 End If
 End Sub
 Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs)
 End Sub
 Private Sub Button5_Click(sender As Object, e As EventArgs) Handles
Button5.Click
 Thread3 = New System.Threading.Thread(AddressOf Close)
 Thread3.Start()
 End Sub
 Public Sub BrowsePathText_TextChanged(sender As Object, e As EventArgs) Handles
BrowsePathText.TextChanged
 End Sub
 Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
 End Sub
 Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
 End Sub
 Private Sub Timer1_Tick_1(sender As Object, e As EventArgs) Handles Timer1.Tick
 End Sub
 Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
 End Sub
 Private Sub Button7_Click(sender As Object, e As EventArgs) Handles
Button7.Click
 ListBox1.Items.Clear()
 SearchText.Clear()
 ProgressBar1.Value = 0
 End Sub
 Private Sub ProgressBar1_Click(sender As Object, e As EventArgs) Handles
ProgressBar1.Click
 End Sub
 Private Sub BackgroundWorker1_DoWork(sender As Object, e As
System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
 If Timer1.Enabled = True Then
 ProgressBar1.Increment(5)
 Else
 ProgressBar1.Value = 100
 End If
 End Sub
 Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As
System.ComponentModel.RunWorkerCompletedEventArgs) Handles
BackgroundWorker1.RunWorkerCompleted
 ProgressBar1.Value = 100
 End Sub
End Class
