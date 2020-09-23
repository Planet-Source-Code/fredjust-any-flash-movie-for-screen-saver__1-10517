VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "swflash.ocx"
Begin VB.Form frmConfig 
   Caption         =   "Settings"
   ClientHeight    =   3570
   ClientLeft      =   3555
   ClientTop       =   2850
   ClientWidth     =   6330
   Icon            =   "Config.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3570
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraQuality 
      Caption         =   "Quality"
      Height          =   1515
      Left            =   60
      TabIndex        =   4
      Tag             =   "High"
      Top             =   1800
      Width           =   1575
      Begin VB.OptionButton Option1 
         Caption         =   "Best"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   675
      End
      Begin VB.OptionButton Option1 
         Caption         =   "High"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   675
      End
      Begin VB.OptionButton Option1 
         Caption         =   "AutoHigh"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "AutoLow"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1035
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Low"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "BackGround"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   780
      Width           =   1515
   End
   Begin VB.CheckBox chkCursor 
      Caption         =   "Show Cursor"
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   1380
      Width           =   1275
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change SWF"
      Height          =   435
      Left            =   120
      TabIndex        =   1
      Top             =   300
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "(ESC for quit Screen Saver)"
      Height          =   435
      Left            =   60
      TabIndex        =   10
      Top             =   3360
      Width           =   1590
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash 
      Height          =   3075
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   4035
      _cx             =   4201421
      _cy             =   4199728
      Movie           =   ""
      Src             =   ""
      WMode           =   "Opaque"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   "000000"
      SWRemote        =   ""
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================================================================
'   Réalisation de Frédéric Just
'   Commentaires remarques et critiques :
'
'   adresse en cours    : fred.just@free.fr
'   site actuel         : http://fred.just.free.fr/
'   adresse de secours  : fredjust@hotmail.com
'==================================================================================

Option Explicit

'===============================================================================
'
'===============================================================================
Private Sub cmdBack_Click()
Dim i As Long

    If VBChooseColor(i) Then
        Flash.BackgroundColor = i
    End If
End Sub

'===============================================================================
'
'===============================================================================
Private Sub cmdChange_Click()
On Error Resume Next
    Dim Tempo As String
    If VBGetOpenFileName(Tempo, "Choose a SWF", , , , , "Macromedia SWF|*.swf") Then
        PlayFlashMovie Tempo
    End If
End Sub

'===============================================================================
'
'===============================================================================
Private Function PlayFlashMovie(Filename As String)
    With Flash
        .Movie = Filename
        .Play
    End With
End Function

'===============================================================================
'
'===============================================================================
Private Sub Command1_Click()
    Flash.StopPlay
End Sub

'===============================================================================
'
'===============================================================================
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

'===============================================================================
'
'===============================================================================
Private Sub Form_Load()
Dim i As Long
On Error Resume Next
    Flash.BackgroundColor = GetSetting("FlashSaver", "Option", "BackColor", 0)
    Flash.Quality2 = GetSetting("FlashSaver", "Option", "Quality", "High")
    PlayFlashMovie GetSetting("FlashSaver", "File", "Anim", App.Path & "\plane.swf")
    
    If CBool(GetSetting("FlashSaver", "Option", "ShowCursor", "False")) Then
        chkCursor.Value = vbChecked
    End If
    
    For i = 0 To 4
        If Option1(i).Caption = GetSetting("FlashSaver", "Option", "Quality", "High") Then
            Option1(i).Value = vbChecked
        End If
    Next
End Sub

'===============================================================================
'
'===============================================================================
Private Sub Form_Resize()
On Error Resume Next
Dim Tempo As String
    Tempo = Flash.Quality2
    Flash.Quality2 = "Low"
    Flash.Move 1800, 120, Me.ScaleWidth - 1800 - 120, Me.ScaleHeight - 120 - 120
    Flash.Quality2 = Tempo
End Sub

'===============================================================================
'
'===============================================================================
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveSetting("FlashSaver", "File", "Anim", Flash.Movie)
    Call SaveSetting("FlashSaver", "Option", "BackColor", Flash.BackgroundColor)
    Call SaveSetting("FlashSaver", "Option", "ShowCursor", CStr(chkCursor.Value))
    Call SaveSetting("FlashSaver", "Option", "Quality", fraQuality.Tag)
End Sub

'===============================================================================
'
'===============================================================================
Private Sub Option1_Click(Index As Integer)
On Error Resume Next
    Flash.Quality2 = Option1(Index).Caption
    fraQuality.Tag = Option1(Index).Caption
End Sub
