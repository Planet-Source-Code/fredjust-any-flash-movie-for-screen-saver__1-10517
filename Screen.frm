VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "swflash.ocx"
Begin VB.Form frmScreen 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6195
   DrawMode        =   4  'Mask Not Pen
   DrawWidth       =   2
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash 
      Height          =   1455
      Left            =   180
      TabIndex        =   0
      Top             =   420
      Width           =   1935
      _cx             =   4197717
      _cy             =   4196870
      Movie           =   ""
      Src             =   ""
      WMode           =   "Opaque"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "Low"
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
Attribute VB_Name = "frmScreen"
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
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

'===============================================================================
'
'===============================================================================
Private Sub Form_Load()

    Flash.BackgroundColor = GetSetting("FlashSaver", "Option", "BackColor", 0)
    Flash.Quality2 = GetSetting("FlashSaver", "Option", "Quality", "High")
    
    With Flash
        .TOp = 0
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight
    End With
    
    PlayFlashMovie GetSetting("FlashSaver", "File", "Anim", App.Path & "\plane.swf")
End Sub

'===============================================================================
'
'===============================================================================
Private Sub Form_Resize()
    With Flash
        .TOp = 0
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight
    End With
'    With lblUp
'        .TOp = 0
'        .Left = 0
'        .Width = Me.ScaleWidth
'        .Height = Me.ScaleHeight
'    End With
End Sub

'===============================================================================
'
'===============================================================================
Private Sub Form_Unload(Cancel As Integer)
    ShowCursor True
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

'Private Sub lblUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Unload Me
'End Sub

'Private Sub lblUp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Static NumMove As Long
'    NumMove = NumMove + 1
'    If NumMove > 10 Then
'        Unload Me
'    End If
'End Sub
