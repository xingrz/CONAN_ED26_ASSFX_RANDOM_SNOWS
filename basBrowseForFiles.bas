Attribute VB_Name = "basBrowseForFiles"
Option Explicit

Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public sFileName() As String '用于打开文件时的数组

Function OpenFiles(ByVal hwnd As Long, ByVal sTitle As String, ByVal sFilter As String, ByVal uFlag As Long, Optional lMaxFileNum As Long) As String
    '说明：
    'hwnd 指窗体的HWND值
    'sTitle 指对话框的标题
    'sTytle 指打开文件的格式，如：*.exe或者*.*，多个可以使用分号隔离
    'uFlag 指是否可以选择多个文件，1可以，其他就只能选一个
    'lMaxFileNum 文件个数，长度的数值，一般是255，如果文件多，可以使用65535

    Dim ofn As OPENFILENAME

    Const OFN_ALLOWMULTISELECT = &H200
    Const OFN_EXPLORER = &H80000
    Const OFN_FILEMUSTEXIST = &H1000

    If Len(Trim(Str(lMaxFileNum))) = 0 Then lMaxFileNum = 255

    uFlag = IIf(uFlag = 1, OFN_ALLOWMULTISELECT Or OFN_EXPLORER Or OFN_FILEMUSTEXIST, OFN_EXPLORER Or OFN_FILEMUSTEXIST)

    With ofn
        .lStructSize = Len(ofn)
        .hwndOwner = hwnd
        .hInstance = App.hInstance
        .lpstrFile = Space$(lMaxFileNum - 1)
        .nMaxFile = lMaxFileNum
        .lpstrFileTitle = Space$(lMaxFileNum - 1)
        .nMaxFileTitle = lMaxFileNum
        .lpstrTitle = sTitle
        .lpstrFilter = CStr(Replace(sFilter, "|", Chr$(0))) '"All Surported Files" + Chr$(0) + sStyle + Chr$(0)
        .flags = uFlag
    End With

    Dim lRet As Long
    lRet = GetOpenFileName(ofn)
    OpenFiles = IIf(lRet > 0, ofn.lpstrFile, "")
    'Debug.Print lRet
End Function

Function GetFileNames(ByVal tmpString As String) As Integer
    Dim FileNames() As String
    FileNames() = Split(tmpString, vbNullChar)
    If UBound(FileNames()) < 3 Then
        ReDim sFileName(0)
        sFileName(0) = FileNames(0) '如果只是选了一个
        GetFileNames = 0
    Else
        Dim m As Integer
        GetFileNames = UBound(FileNames) - 3
        ReDim sFileName(0 To GetFileNames)
        For m = 0 To GetFileNames
            sFileName(m) = IIf(Right(FileNames(0), 1) = "\", FileNames(0) + FileNames(m + 1), FileNames(0) + "\" + FileNames(m + 1))
        Next
    End If
End Function
