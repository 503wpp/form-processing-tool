VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000004&
   Caption         =   "待办数据与台账数据对比"
   ClientHeight    =   7470
   ClientLeft      =   3585
   ClientTop       =   705
   ClientWidth     =   14535
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   14535
   Begin VB.Frame Frame4 
      Caption         =   "Frame4"
      Height          =   975
      Left            =   6120
      TabIndex        =   14
      Top             =   7560
      Width           =   2415
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   1215
      Left            =   3120
      TabIndex        =   12
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1095
      Left            =   720
      TabIndex        =   2
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   6495
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   13815
      Begin VB.CommandButton Command1 
         Caption         =   "PSM流程汇总"
         Height          =   495
         Left            =   600
         TabIndex        =   13
         Top             =   3300
         Width           =   1575
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   240
         Top             =   4200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3855
         Left            =   4560
         TabIndex        =   10
         Top             =   795
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   6800
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   450
         Left            =   600
         TabIndex        =   9
         Top             =   5115
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   794
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton Command5 
         Caption         =   "导出差异数据"
         Height          =   495
         Left            =   2520
         TabIndex        =   8
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "导出全部数据"
         Height          =   495
         Left            =   2520
         TabIndex        =   7
         Top             =   1500
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "开始对比"
         Height          =   495
         Left            =   600
         TabIndex        =   6
         Top             =   1500
         Width           =   1575
      End
      Begin VB.CommandButton Command7 
         Caption         =   "导出待办数据"
         Height          =   495
         Left            =   2520
         TabIndex        =   5
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         Caption         =   "开始拆分"
         Height          =   495
         Left            =   600
         TabIndex        =   4
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "对比结果"
         Height          =   195
         Left            =   4560
         TabIndex        =   11
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "运行进度"
         Height          =   195
         Left            =   600
         TabIndex        =   3
         Top             =   4920
         Width           =   6135
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   13150
      MultiRow        =   -1  'True
      TabStyle        =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "待办数据处理"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "小工具2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "小工具3"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "帮助"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sourceFilePath_1 As String
Public sourceFilePath_2 As String
Public sourceFilePath_3 As String
Public ProgressValue As Integer

Sub Command1_Click()
    Call 选择PSM进度汇总模板路径.选择PSM进度汇总模板路径
    
    Call 开始汇总.开始汇总
End Sub

Sub Command3_Click()
    Call 选择台账数据路径.选择台账数据路径
    Call 开始对比.开始对比
End Sub

Sub Command4_Click()
    Call 导出全部数据.导出全部数据
End Sub

Sub Command5_Click()
    Call 导出差异数据.导出差异数据
End Sub

Sub Command6_Click()
    Call 选择待办数据路径.选择待办数据路径
    Call 开始拆分.开始拆分
End Sub

Sub Command7_Click()
    Call 导出待办数据.导出待办数据
End Sub

Private Sub Form_Load()
    Dim frm As Frame
    For i = 1 To TabStrip1.Tabs.Count
        Form1.Controls("Frame" & i).Width = TabStrip1.ClientWidth 'tab内部宽度
        Form1.Controls("Frame" & i).Height = TabStrip1.ClientHeight 'tab内部高度
        Form1.Controls("Frame" & i).Left = TabStrip1.ClientLeft 'tab内部左间距
        Form1.Controls("Frame" & i).Top = TabStrip1.ClientTop 'tab内部顶部间距
    Next i
    
    For i = 2 To TabStrip1.Tabs.Count
        Form1.Controls("Frame" & i).Visible = False
    Next i
    
End Sub


Private Sub TabStrip1_Click()
    For i = 1 To TabStrip1.Tabs.Count
    If TabStrip1.SelectedItem.Index = i Then
        Form1.Controls("Frame" & i).Tag = "dq"
    End If
    If Form1.Controls("Frame" & i).Tag = "dq" Then
        Form1.Controls("Frame" & i).Visible = True
    Else
        Form1.Controls("Frame" & i).Visible = False
    End If
    Form1.Controls("Frame" & i).Tag = Empty
    Next i

End Sub


