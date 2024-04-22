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

Public sFileName() As String '���ڴ��ļ�ʱ������

Function OpenFiles(ByVal hwnd As Long, ByVal sTitle As String, ByVal sFilter As String, ByVal uFlag As Long, Optional lMaxFileNum As Long) As String
    '˵����
    'hwnd ָ�����HWNDֵ
    'sTitle ָ�Ի���ı���
    'sTytle ָ���ļ��ĸ�ʽ���磺*.exe����*.*���������ʹ�÷ֺŸ���
    'uFlag ָ�Ƿ����ѡ�����ļ���1���ԣ�������ֻ��ѡһ��
    'lMaxFileNum �ļ����������ȵ���ֵ��һ����255������ļ��࣬����ʹ��65535

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
        sFileName(0) = FileNames(0) '���ֻ��ѡ��һ��
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
