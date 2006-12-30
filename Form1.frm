VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ED26 仿雪花随机摆动特效生成"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   4680
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdBrowers 
      Caption         =   "..."
      Height          =   255
      Left            =   4080
      TabIndex        =   17
      Top             =   240
      Width           =   495
   End
   Begin VB.TextBox txtFile 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1080
      TabIndex        =   16
      Top             =   240
      Width           =   2895
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出"
      Height          =   375
      Left            =   3360
      TabIndex        =   12
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "备份恢复"
      Height          =   2295
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   4455
      Begin VB.CommandButton cmdBackupDel 
         Caption         =   "删除"
         Height          =   375
         Left            =   1200
         TabIndex        =   19
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton cmdBackupRef 
         Caption         =   "刷新"
         Height          =   375
         Left            =   2280
         TabIndex        =   18
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton cmdRecover 
         Caption         =   "恢复"
         Height          =   375
         Left            =   3360
         TabIndex        =   8
         Top             =   1800
         Width           =   975
      End
      Begin VB.ListBox lstBackup 
         Height          =   1500
         ItemData        =   "Form1.frx":0000
         Left            =   120
         List            =   "Form1.frx":0002
         TabIndex        =   7
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "特效生成"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   4455
      Begin VB.TextBox txtYrate 
         Height          =   270
         Left            =   1320
         TabIndex        =   10
         Text            =   "10"
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox chkBackup 
         Caption         =   "创建备份"
         Height          =   180
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.TextBox txtXend 
         Height          =   270
         Left            =   2040
         TabIndex        =   4
         Text            =   "0"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtXbegin 
         Height          =   270
         Left            =   1320
         TabIndex        =   3
         Text            =   "0"
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmdMake 
         Caption         =   "生成"
         Height          =   375
         Left            =   3360
         TabIndex        =   1
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "像素/秒"
         Height          =   180
         Left            =   1920
         TabIndex        =   11
         Top             =   720
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Y 轴速度:"
         Height          =   180
         Left            =   360
         TabIndex        =   9
         Top             =   720
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "X 轴幅度:"
         Height          =   180
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   810
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "选择文件:"
      Height          =   180
      Left            =   120
      TabIndex        =   15
      Top             =   240
      Width           =   810
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "星光字幕组 www.conans.com"
      Height          =   180
      Left            =   240
      TabIndex        =   14
      Top             =   5040
      Width           =   2250
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "制作:chenxingyu"
      Height          =   180
      Left            =   240
      TabIndex        =   13
      Top             =   4800
      Width           =   1350
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objFSO As Object

Private Sub cmdBackupDel_Click()
    Dim i As Integer
    For i = 0 To lstBackup.ListCount - 1
        If lstBackup.Selected(i) Then
            Call objFSO.DeleteFile(txtFile.Text & "." & CntShotDate(Trim(lstBackup.List(i))) & ".backup")
            Exit For
        End If
    Next
    Call RescanBackups
End Sub

Private Sub cmdBackupRef_Click()
    Call RescanBackups
End Sub

Private Sub cmdBrowers_Click()
    Dim objFile As Object
    Dim tmp
    Dim prx As Integer
    tmp = OpenFiles(Me.hwnd, "请选择一个文件", "ASS 特效字幕文件 (*.ass)|*.ass", 0, 255)
    If Trim(tmp) = "" Then Exit Sub
    txtFile.Text = tmp
    Set objFile = objFSO.OpenTextFile(txtFile.Text, 1)
    Do While Not objFile.AtEndOfStream
        tmp = objFile.ReadLine
        If Left(tmp, 10) = "PlayResX: " Then prx = CInt(Mid(tmp, 11)): Exit Do
    Loop
    Set objFile = Nothing
    If prx > 0 Then
    txtXbegin.Text = CStr(prx - 60)
    txtXend.Text = CStr(prx - 20)
    End If
    Call RescanBackups
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdMake_Click()
    If chkBackup.Value = vbChecked Then
        Call objFSO.CopyFile(txtFile.Text, txtFile.Text & "." & GetShotDate & ".backup")
        Call RescanBackups
    End If
    Call 生成特效
End Sub

Private Sub cmdRecover_Click()
    Dim i As Integer
    For i = 0 To lstBackup.ListCount - 1
        If lstBackup.Selected(i) Then
            Call objFSO.CopyFile(txtFile.Text & "." & CntShotDate(Trim(lstBackup.List(i))) & ".backup", txtFile.Text, True)
            Call objFSO.DeleteFile(txtFile.Text & "." & CntShotDate(Trim(lstBackup.List(i))) & ".backup")
            Exit For
        End If
    Next
    Call RescanBackups
End Sub

Private Sub Form_Load()
    Set objFSO = CreateObject("Scripting.FileSystemObject")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objFSO = Nothing
End Sub

Private Sub 生成特效()
    '######################################################
    '#       名侦探柯南 ED26 仿雪花随机摆动特效生成       #
    '#                                                    #
    '# 制作:chenxingyu                                    #
    '# 星光字幕组 www.conans.com                          #
    '#                                                    #
    '######################################################
    Dim 摆动范围 As Integer
    Dim 替换关键词 As String
    Dim objFile As Object
    Dim tmp
    Dim 等待写入 As String
    摆动范围 = CInt(txtXend.Text) - CInt(txtXbegin.Text)
    Set objFile = objFSO.OpenTextFile(txtFile.Text, 1)
    Do While Not objFile.AtEndOfStream
        tmp = objFile.ReadLine
        If InStr(tmp, "Dialogue: ") > 0 Then
            Dim tmp2, tmp3
            Dim i As Integer
            Dim 横始坐标, 横末坐标 As Integer
            Dim 纵始坐标, 纵末坐标 As Double
            Dim 开始时间, 结束时间 As Double
            横始坐标 = CInt(txtXbegin.Text) + Int(Rnd() * 摆动范围)
            横末坐标 = CInt(txtXbegin.Text) + Int(Rnd() * 摆动范围)
            tmp2 = Split(tmp, ",")
            开始时间 = GetTimems(tmp2(1))
            结束时间 = GetTimems(tmp2(2))
            tmp3 = Split(tmp, "\")
            For i = 0 To UBound(tmp3) - 1
                'Debug.Print tmp3(i)
                If Left(tmp3(i), 2) = "vb" Then
                    纵始坐标 = CInt(Mid(tmp3(i), 3))
                    纵末坐标 = CInt(纵始坐标 + ((结束时间 - 开始时间) / 1000 * CInt(txtYrate.Text)))
                    tmp = Replace(tmp, tmp3(i), "move(" & 横始坐标 & "," & 纵始坐标 & "," & 横末坐标 & "," & 纵末坐标 & ")")
                End If
            Next
        End If
        等待写入 = 等待写入 & tmp & vbNewLine
    Loop
    Set objFile = Nothing
    Set objFile = objFSO.OpenTextFile(txtFile.Text, 2)
    objFile.Write 等待写入
    Set objFile = Nothing
    MsgBox "完成！"
End Sub

Private Function GetTimems(ByVal Timestamp As String) As Double
    Dim h, m, s, ms As Integer
    h = CInt(Mid(Timestamp, 1, 1))
    m = CInt(Mid(Timestamp, 3, 2))
    s = CInt(Mid(Timestamp, 6, 2))
    ms = CInt(Mid(Timestamp, 9, 2))
    GetTimems = h * 60 * 60 * 1000 + m * 60 * 1000 + s * 1000 + ms * 10
End Function

Private Function GetShotDate() As String
    GetShotDate = _
    Year(Now) & _
    IIf(Len(Month(Now)) < 2, "0" & Month(Now), Month(Now)) & _
    IIf(Len(Day(Now)) < 2, "0" & Day(Now), Day(Now)) & _
    IIf(Len(Hour(Now)) < 2, "0" & Hour(Now), Hour(Now)) & _
    IIf(Len(Minute(Now)) < 2, "0" & Minute(Now), Minute(Now)) & _
    IIf(Len(Second(Now)) < 2, "0" & Second(Now), Second(Now))
End Function

Private Function CntFullDate(ByVal ShotDate As String) As String
    CntFullDate = _
    Mid(ShotDate, 1, 4) & "-" & Mid(ShotDate, 5, 2) & "-" & Mid(ShotDate, 7, 2) & " " & _
    Mid(ShotDate, 9, 2) & ":" & Mid(ShotDate, 11, 2) & ":" & Mid(ShotDate, 13, 2)
End Function

Private Function CntShotDate(ByVal FullDate As String) As String
    CntShotDate = _
    Mid(FullDate, 1, 4) & _
    Mid(FullDate, 6, 2) & _
    Mid(FullDate, 9, 2) & _
    Mid(FullDate, 12, 2) & _
    Mid(FullDate, 15, 2) & _
    Mid(FullDate, 18, 2)
End Function

Private Sub RescanBackups()
    Dim objFile, objFolder, objFiles As Object
    Dim tmp
    If Trim(txtFile.Text) = "" Then Exit Sub
    lstBackup.Clear
    Set objFile = objFSO.GetFile(txtFile.Text)
    Set objFolder = objFile.ParentFolder
    For Each objFiles In objFolder.Files
        tmp = objFiles.Name
        If Right(tmp, 7) = ".backup" Then
            tmp = Split(tmp, ".")
            lstBackup.AddItem (CntFullDate(tmp(UBound(tmp) - 1)))
        End If
    Next
End Sub
