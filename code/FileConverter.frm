VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FileConverter 
   Caption         =   "Convert Excel or Word Files"
   ClientHeight    =   6156
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3648
   OleObjectBlob   =   "FileConverter.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FileConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Convert_Click()
Dim vPath As String
vPath = Me.Path.Text
If Strings.Left(vPath, 1) = """" Then vPath = Mid(vPath, 2)
If Right(vPath, 1) = """" Then vPath = Strings.Left(vPath, Len(vPath) - 1)
If vPath = "" Then Exit Sub
Select Case IsFileOrFolder(vPath)
    Case "Folder"
        If Right(vPath, 1) <> "\" Then vPath = vPath & "\"
        ConvertFiles vPath
    Case "File"
        convertFile vPath
    Case Else
        MsgBox "File or Folder not found. Please check again."
End Select

End Sub

Sub ConvertFiles(vPath As String)
    WorkFile = Dir(vPath & IIf(Me.oExcelFiles.Value = True, "*.xls*", "*.doc*"))
    Do While WorkFile <> ""
        If Right(WorkFile, 4) <> IIf(Me.oExcelFiles.Value = True, "xlsm", "docm") Then
                convertFile vPath & WorkFile
        End If
        WorkFile = Dir()
    Loop
End Sub

Sub convertFile(vPath As String)
    If oExcelFiles.Value = True Then
        Select Case UCase(whichOption(Me.ExcelOutput, "OptionButton").Caption)
            Case "XLSB"
                XLS_ConvertFileFormat vPath, xlExcel12, Me.oDelete
            Case "XLSM"
                XLS_ConvertFileFormat vPath, xlOpenXMLWorkbookMacroEnabled, Me.oDelete
            Case "XLSX"
                XLS_ConvertFileFormat vPath, xlWorkbookDefault, Me.oDelete
            Case "CSV"
                XLS_ConvertFileFormat vPath, xlCSV, Me.oDelete
            Case "XLAM"
                XLS_ConvertFileFormat vPath, xlOpenXMLAddIn, Me.oDelete
            Case "PDF"
                ExcelToPDF vPath, cSeparateSheets.Value, True
        End Select
    Else
        Select Case whichOption(Me.WordOutput, "OptionButton").Caption
            Case "DOCX"
                Word_ConvertFileFormat vPath, wdFormatDocumentDefault, Me.oDelete
            Case "TXT"
                Word_ConvertFileFormat vPath, wdFormatText, Me.oDelete
            Case "DOCM"
                Word_ConvertFileFormat vPath, wdFormatXMLDocumentMacroEnabled, Me.oDelete
            Case "PDF"
                Word_ConvertFileFormat vPath, wdFormatPDF, Me.oDelete
        End Select
    End If
End Sub





Private Sub info_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
uDEV.Show
End Sub

Private Sub oExcelFiles_Click()
WordOutput.Visible = False
ExcelOutput.Visible = True
End Sub
Private Sub oWordFiles_Click()
WordOutput.Visible = True
ExcelOutput.Visible = False
End Sub

Private Sub PickFile_Click()
Dim vPath As String: vPath = SelectFile
If vPath <> "" Then Me.Path.Text = vPath
End Sub
Private Sub PickFolder_Click()
Dim vPath As String: vPath = SelectFolder
If vPath <> "" Then Me.Path.Text = vPath & "\"
End Sub
Function SelectFile() As String
Dim strFile As String
Dim fd As Office.FileDialog
Set fd = Application.FileDialog(msoFileDialogFilePicker)
With fd
    .Filters.Clear
    .Filters.Add IIf(Me.oExcelFiles.Value = True, "Excel Files", "Word Files"), IIf(Me.oExcelFiles.Value = True, "*.xl*", "*.doc*"), 1
    .Title = IIf(Me.oExcelFiles.Value = True, "Choose an Excel file", "Choose a Word File")
    .AllowMultiSelect = False
    .InitialFileName = Environ("USERprofile") & "\Desktop\"
    If .Show = True Then
        strFile = .SelectedItems(1)
        SelectFile = strFile
    End If
End With
End Function
