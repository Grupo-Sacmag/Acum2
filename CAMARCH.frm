VERSION 5.00
Begin VB.Form Camdir 
   Caption         =   "Cambio de Subdirectorio"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7410
   Icon            =   "CAMARCH.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   7410
   StartUpPosition =   2  'CenterScreen
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   4560
      Width           =   2895
   End
   Begin VB.FileListBox File1 
      Height          =   4770
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   3975
   End
   Begin VB.DirListBox Dir1 
      Height          =   4140
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   5160
      Width           =   5535
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Directorio :"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   5160
      Width           =   1215
   End
End
Attribute VB_Name = "Camdir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dir

Private Sub Dir1_Change()

   File1_Click
   File1_DblClick
   
End Sub
Private Sub Dir1_KeyPress(KeyAscii As Integer)
    File1_Click
    File1_DblClick
  
End Sub

Private Sub Drive1_Change()
   On Error GoTo manejodrive
   ChDrive Drive1.Drive
   Dir1.Path = Drive1.Drive
   Dir1 = Dir1.Path
   Dir1_Change
   GoTo saledriv
manejodrive:
   Drive1.Drive = "C:"
   Dir1 = "C:\"
   
saledriv:
End Sub

Private Sub File1_Click()
   If Dir1.Path <> Dir1.List(Dir1.ListIndex) Then
        Dir1.Path = Dir1.List(Dir1.ListIndex)
        File1 = Dir1.Path
        Exit Sub
   End If
   File1 = Dir1.Path
End Sub

Private Sub File1_DblClick()
   File1 = Dir1.Path
   ChDir CurDir(Dir1)
   Label2.Caption = Dir1.Path
End Sub

Private Sub Form_Load()
    
     Label2.Caption = Dir1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Close 3
    MsgBox (Dir1)
    Open "C:\GconTa\sccontr.soc" For Random As 3 Len = Len(SCont)
    SCont.guarda = Dir1
    Put 3, 1, SCont
    Close 3

End Sub
