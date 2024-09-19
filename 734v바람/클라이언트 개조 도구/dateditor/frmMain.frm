VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "@���ٿ� - Baram *Dat Editor v1"
   ClientHeight    =   3255
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   4470
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4470
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog Cd 
      Left            =   3960
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�߰�"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   2760
      Width           =   615
   End
   Begin MSComctlLib.ListView lstDat 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "�̸�"
         Object.Width           =   3351
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ũ��"
         Object.Width           =   3175
      EndProperty
   End
   Begin VB.Menu mnu_hwnd 
      Caption         =   "�޴�"
      Begin VB.Menu mnu_New 
         Caption         =   "�� ������(&New)"
      End
      Begin VB.Menu mnu_Open 
         Caption         =   "�ҷ�����(&Open)"
      End
      Begin VB.Menu mnu_Save 
         Caption         =   "�����ϱ�(&Save)"
      End
      Begin VB.Menu mnu_blank 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnu_Exit 
         Caption         =   "����(&Exit)"
      End
   End
   Begin VB.Menu mnu_About 
      Caption         =   "����"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    On Error GoTo Err:
    Cd.DialogTitle = "�߰��� ������ ������ �ּ���!"
    Cd.Filter = "��� ���� (*.*)|*.*"
    Cd.FileName = ""
    Cd.ShowOpen
    If Trim(Cd.FileName) = "" Then Exit Sub
    AddDat Cd.FileName
Err:
End Sub

Private Sub Command2_Click()
    On Error GoTo Err:
    RemoveDat (lstDat.SelectedItem.Index - 1)
Err:
End Sub

Private Sub Command3_Click()
On Error GoTo Err:
    Cd.DialogTitle = "����� ��ġ�� ������ �ּ���!"
    Cd.Filter = "��� ���� (*.*)|*.*"
    Cd.FileName = lstDat.SelectedItem.Text
    Cd.ShowSave
    If Trim(Cd.FileName) = "" Then Exit Sub
    ExtractDat (lstDat.SelectedItem.Index - 1), Cd.FileName
Err:
End Sub

Private Sub Form_Load()
    ClearDat
End Sub

Private Sub mnu_About_Click()
    Call MsgBox("@���ٿ�!! ��" & vbCrLf & "http://FBStyle.wo.tc" & vbCrLf & "����: ����(love947345) > _<��     ", vbApplicationModal + vbInformation, "����")
End Sub

Private Sub mnu_New_Click()
    ClearDat
End Sub

Private Sub mnu_Open_Click()
On Error GoTo Err:
    Cd.DialogTitle = "�ҷ��� Dat������ ������ �ּ���!"
    Cd.Filter = "�ٶ��� ���� Dat���� (*.dat)|*.dat"
    Cd.ShowOpen
    If Trim(Cd.FileName) = "" Then Exit Sub
    OpenDat Cd.FileName
Err:
End Sub

Private Sub mnu_Save_Click()
On Error GoTo Err:
    Cd.DialogTitle = "����� Dat������ �Է��� �ּ���!"
    Cd.Filter = "�ٶ��� ���� Dat���� (*.dat)|*.dat"
    Cd.ShowSave
    If Trim(Cd.FileName) = "" Then Exit Sub
    SaveDat Cd.FileName
Err:
End Sub
