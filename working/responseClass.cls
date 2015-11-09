VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "responseClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private folderID As Long

Private CSV As Object
Private Main As Object

Public Function Execute(CsvObject As Object, MainObject As Object, OptionString As String, FilterInput As String) As String
    Set CSV = CsvObject
    
    Call Init(MainObject)
    
    Execute = GetContent(OptionString)
    
End Function

Public Sub Init(MainObject As Object)
    Set Main = MainObject
    
    Exit Sub
ErrorTrap:
    Call HandleError("responseClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
End Sub

Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap

    Dim sizeCnt As Long
    Dim Ptr As Long
    Dim sourceFile As String
    Dim altSizeList As String
    Dim sizes() As String
    Dim Pos As Long
    Dim dimensions() As String
    Dim imageSrc As String
    Dim thumbSrc As String
    Dim imageName As String
    Dim CS As Long
    Dim CS2 As Long
    Dim Stream As String
    Dim folderID As Long
    Dim folderName As String
    Dim baseFile As String
    Dim fileType As String
    Dim imageFile As String
    Dim thumbFile As String
    Dim objFile As New kmaFileSystem3.FileSystemClass
    Dim imagesNode As String

    Stream = Main.ReadBake("photoGallery")

    If Stream = "" Then
        
        Stream = "" _
            & CR & "<options>" _
            & kmaIndent(CR & "<spacing categoryMargin=""0"" categoryPadding=""5,5,5,5"" thumbMargin=""5"" thumbPadding=""5"" />") _
            & CR & "</options>"
        'Stream = Stream & "<backgrounds category=""000000"" categoryActive=""000000"" thumb=""000000"" thumbActive=""000000"" photoBackground=""000000"" />" & vbCrLf
        'Stream = Stream & "<text category=""000000"" categoryActive=""000000"" thumb=""000000"" thumbActive=""000000"" />" & vbCrLf
        'Stream = Stream & "<borders category=""#000000"" thumb=""#000000"" photo=""#000000"" />" & vbCrLf
        'Stream = Stream & "<icons categoryIcon=""images/icon.gif"" />" & vbCrLf
        '
        folderID = Main.GetRecordID("Library Folders", "Photo Gallery")
        If folderID = 0 Then
            CS = Main.InsertCSContent("Library Folders")
            If Main.CSOK(CS) Then
                folderID = Main.GetCSInteger(CS, "ID")
                Call Main.SetCS(CS, "Name", "Photo Gallery")
            End If
            Call Main.CloseCS(CS)
        End If
        '
        imagesNode = getFolderImages(folderID, "Photo Gallery")
        CS = Main.OpenCSContent("Library Folders", "ParentID=" & folderID, "SortOrder")
        If Main.IsCSOK(CS) Then
            Do While Main.CSOK(CS)
                folderID = Main.GetCSInteger(CS, "ID")
                folderName = Main.GetCSText(CS, "name")
                imagesNode = imagesNode & getFolderImages(folderID, folderName)
                Call Main.NextCSRecord(CS)
            Loop
        End If
        If imagesNode <> "" Then
            Stream = "" _
                & Stream _
                & CR & "<images>" _
                & kmaIndent(imagesNode) _
                & CR & "</images>"
        End If
        Call Main.CloseCS(CS)
        '
        Stream = "" _
            & CR & "<gallery>" _
            & kmaIndent(Stream) _
            & CR & "</gallery>"
        
        Call Main.SaveBake("photoGallery", Stream, "Library Files")
        
    End If
    
    GetContent = Stream
    
    Exit Function
ErrorTrap:
    Call HandleError("responseClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
End Function
'
'
'
Public Function getFolderImages(folderID As Long, folderName As String) As String
    On Error GoTo ErrorTrap

    Dim sizeCnt As Long
    Dim Ptr As Long
    Dim sourceFile As String
    Dim altSizeList As String
    Dim sizes() As String
    Dim Pos As Long
    Dim dimensions() As String
    Dim imageSrc As String
    Dim thumbSrc As String
    Dim imageName As String
    Dim CS As Long
    Dim CS2 As Long
    Dim Stream As String
    Dim baseFile As String
    Dim fileType As String
    Dim imageFile As String
    Dim thumbFile As String
    Dim objFile As New kmaFileSystem3.FileSystemClass

    CS2 = Main.OpenCSContent("Library Files", "FolderID=" & folderID)
    Do While Main.CSOK(CS2)
        sourceFile = Main.GetCSText(CS2, "FileName")
        altSizeList = Main.GetCSText(CS2, "altSizeList")
        imageName = Main.GetCSText(CS2, "Name")
        '
        sourceFile = Replace(sourceFile, "/", "\")
        baseFile = sourceFile
        Pos = InStrRev(baseFile, ".", , vbTextCompare)
        If Pos > 0 Then
            fileType = Mid(baseFile, Pos + 1)
            baseFile = Left(baseFile, Pos - 1)
        End If
        '
        imageFile = sourceFile
        thumbFile = sourceFile
        If altSizeList <> "" Then
            sizes = Split(altSizeList, vbCrLf)
            sizeCnt = UBound(sizes) + 1
            For Ptr = 0 To sizeCnt - 1
                dimensions = Split(sizes(Ptr), "x")
                If UBound(dimensions) > 0 Then
                    If dimensions(0) = "640" Then
                        imageFile = baseFile & "-" & sizes(Ptr) & "." & fileType
                    ElseIf dimensions(0) = "160" Then
                        thumbFile = baseFile & "-" & sizes(Ptr) & "." & fileType
                    End If
                End If
            Next
        End If
            
        'Call AppendLogFile("photoGallery response thumbbnail [" & thumbFile & "], altSizeList [" & altSizeList & "]")
        
        If InStr(1, thumbFile, "x", vbTextCompare) = 0 Then
            Call AppendLogFile("photoGallery response did not find a -160x version of the image [" & Main.PhysicalFilePath & imageFile & "]")
        End If
        If Not objFile.CheckFile(Main.PhysicalFilePath & thumbFile) Then
            Call AppendLogFile("photoGallery response failed to find the thumbnail image [" & Main.PhysicalFilePath & thumbFile & "]")
            thumbFile = sourceFile
        End If
        'If Not objFile.CheckFile(Main.PhysicalFilePath & imageFile) Then
            imageFile = sourceFile
        'End If
        '
        imageSrc = Replace(imageFile, "\", "/")
        imageSrc = kmaEncodeURL(imageSrc)
        '
        thumbSrc = Replace(thumbFile, "\", "/")
        thumbSrc = kmaEncodeURL(thumbSrc)
        '
        Stream = Stream & CR & "<image large=""" & kmaEncodeHTML(imageSrc) & """ thumb=""" & kmaEncodeHTML(thumbSrc) & """ title=""" & kmaEncodeHTML(imageName) & """ />"
        Call Main.NextCSRecord(CS2)
    Loop
    Call Main.CloseCS(CS2)
    
    If Stream <> "" Then
        Stream = "" _
            & CR & "<category title=""" & kmaEncodeHTML(folderName) & """ folder=""" & Main.ServerFilePath & """>" _
            & kmaIndent(Stream) _
            & CR & "</category>"
    End If
    
    getFolderImages = Stream
    
    Exit Function
ErrorTrap:
    Call HandleError("responseClass", "getFolderImages", Err.Number, Err.Source, Err.Description, True, False)
End Function

