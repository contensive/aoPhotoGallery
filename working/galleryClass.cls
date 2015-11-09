VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "galleryClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Main As Object
Private CSV As Object

Public Function Execute(CsvObject As Object, MainObject As Object, OptionString As String, FilterInput As String) As String
    Set CSV = CsvObject
    
    Call Init(MainObject)
    
    Execute = GetContent(OptionString)
    
End Function

Public Sub Init(MainObject As Object)
    Set Main = MainObject
    
    Exit Sub

ErrorTrap:
    Call HandleError("galleryClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
    End Sub

Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    
    Dim Stream As String
    Dim locWidth As Long
    Dim locHeight As Long
    
    locWidth = kmaEncodeInteger(Main.GetAddonOption("Photo Gallery Width", OptionString))
    If locWidth < 1 Then
        locWidth = 650
    End If
    
    locHeight = (450 * locWidth) / 650
    'locWidth = Main.GetSiteProperty("Photo Gallery Width", "650")
    'locHeight = Main.GetAddonOption("Photo Gallery Height", OptionString)
    'locHeight = Main.GetSiteProperty("Photo Gallery Height", "450")

    Stream = Stream & "<object classid=""clsid:d27cdb6e-ae6d-11cf-96b8-444553540000"" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,0,0"" width=""" & locWidth & """ height=""" & locHeight & """ id=""photo_gallery"" align=""middle"">"
    Stream = Stream & "<param name=""allowScriptAccess"" value=""sameDomain"" />"
    Stream = Stream & "<param name=""allowFullScreen"" value=""false"" />"
    Stream = Stream & "<param name=""movie"" value=""/photo_gallery1.swf"" />"
    Stream = Stream & "<param name=""quality"" value=""high"" />"
    Stream = Stream & "<param name=""bgcolor"" value=""#ffffff"" />"
    Stream = Stream & "<param name=""wmode"" value=""transparent"" />"
    Stream = Stream & "<embed src=""/photo_gallery1.swf"" wmode=""transparent"" quality=""high"" bgcolor=""#ffffff"" width=""" & locWidth & """ height=""" & locHeight & """ name=""photo_gallery1"" align=""middle"" allowScriptAccess=""sameDomain"" allowFullScreen=""false"" type=""application/x-shockwave-flash"" pluginspage=""http://www.adobe.com/go/getflashplayer"" />"
    Stream = Stream & "</object>"
    '
    If Main.IsAdmin() Then
        Stream = Stream & Main.GetAdminHintWrapper("<p>Add photos and folders to the Photo Gallery by added them to the <a href=""" & Main.SiteProperty_AdminURL & "?addonguid=%7B564EF3F5-9673-4212-A692-0942DD51FF1A%7D"" target=""_blank"">Resource Library</a> in the Photo Gallery Folder</p>")
    End If
    
    GetContent = Stream
    
    Exit Function
ErrorTrap:
    Call HandleError("galleryClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function

