VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmExplorerV2 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Visor de tickets"
   ClientHeight    =   11820
   ClientLeft      =   120
   ClientTop       =   390
   ClientWidth     =   9195
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExplorerV2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11820
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtZoom 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "100"
      Top             =   135
      Width           =   1335
   End
   Begin SHDocVwCtl.WebBrowser wbExplorer 
      CausesValidation=   0   'False
      Height          =   10695
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   960
      Width           =   8895
      ExtentX         =   15690
      ExtentY         =   18865
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Image imgCursor 
      Height          =   480
      Left            =   0
      Picture         =   "frmExplorerV2.frx":08CA
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMas 
      Height          =   480
      Left            =   2880
      Picture         =   "frmExplorerV2.frx":1194
      ToolTipText     =   "Salir"
      Top             =   120
      Width           =   480
   End
   Begin VB.Image imgMenos 
      Height          =   480
      Left            =   600
      Picture         =   "frmExplorerV2.frx":1A5E
      ToolTipText     =   "Salir"
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "* Click sobre la imagen para rotar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   3345
   End
End
Attribute VB_Name = "frmExplorerV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim msRuta As String
Dim miAplicacion As Integer

Private Const OLECMDID_OPTICAL_ZOOM As Long = 63
Private Const OLECMDEXECOPT_DONTPROMPTUSER As Long = 2

Private mblHasSetZoom As Boolean

Public Function Create(sRuta As String, Optional iAplicacion As Integer = 0)
    msRuta = sRuta
    miAplicacion = iAplicacion
End Function

Private Sub Form_Activate()

    txtZoom = 100
    If msRuta = "" Then Exit Sub
    Select Case miAplicacion
    Case 0
        Me.Caption = "Visor de tickets"
    Case 1
        Me.Caption = "Visor de imágenes"
    Case Else
        Me.Caption = "Visor"
    End Select
    wbExplorer.Navigate msRuta

End Sub

Private Sub Form_Load()
    
'    If DisplayMonitorCount = 1 Then
'        Me.Left = (Screen.Width / 2) - (Me.Width / 2)
'        Me.Top = (Screen.Height / 2) - (Me.Height / 2)
'    Else
'        If MDIPrincipal.Left + 120 < Screen.Width Then
'        'El MDI está en la pantalla 1, lo centro en la pantalla 1
'            Me.Left = (Screen.Width / 2) - (Me.Width / 2)
'            Me.Top = (Screen.Height / 2) - (Me.Height / 2)
'        Else
'            Me.Left = gAnchoPantallaPrincipal + (((gAnchoPantallaVirtual - gAnchoPantallaPrincipal) - Me.Width) \ 2)
'            Me.Top = (Screen.Height - Me.Height) \ 2
'        End If
'    End If
    
    txtZoom = 100
    If msRuta = "" Then Exit Sub
    Select Case miAplicacion
    Case 0
        Me.Caption = "Visor de tickets"
    Case 1
        Me.Caption = "Visor de imágenes"
    Case Else
        Me.Caption = "Visor"
    End Select
    wbExplorer.Navigate msRuta
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = DEFAULT
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        'MsgBox "Si quiere cerrar la ventana, pulse el botón 'Salir'", vbApplicationModal + vbExclamation, "Información"
        Me.Hide
        Cancel = True
    End If
End Sub

Private Sub Form_Resize()
    wbExplorer.Width = Me.ScaleWidth - 330
    wbExplorer.Height = Me.ScaleHeight - 1125
End Sub

Private Sub imgMas_Click()
    Me.MousePointer = DEFAULT
    If txtZoom = 200 Then Exit Sub
    txtZoom = txtZoom + 10
    Zoom txtZoom
End Sub

Private Sub imgMas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgMenos_Click()
    Me.MousePointer = DEFAULT
    If txtZoom = 30 Then Exit Sub
    txtZoom = txtZoom - 10
    Zoom txtZoom
End Sub

Private Sub imgMenos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub imgSalir_Click()
    Me.MousePointer = DEFAULT
    Unload Me
End Sub

Private Sub imgSalir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Me.MouseIcon = imgCursor.Picture
    Me.MousePointer = 99
End Sub

Private Sub Zoom(iZoom As Integer)

    Select Case iZoom
    Case 30
        wbExplorer.ExecWB OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, 30&
    Case 40
        wbExplorer.ExecWB OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, 40&
    Case 50
        wbExplorer.ExecWB OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, 50&
    Case 60
        wbExplorer.ExecWB OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, 60&
    Case 70
        wbExplorer.ExecWB OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, 70&
    Case 80
        wbExplorer.ExecWB OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, 80&
    Case 90
        wbExplorer.ExecWB OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, 90&
    Case 100
        wbExplorer.ExecWB OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, 100&
    Case 110
        wbExplorer.ExecWB OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, 110&
    Case 120
        wbExplorer.ExecWB OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, 120&
    Case 130
        wbExplorer.ExecWB OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, 130&
    Case 140
        wbExplorer.ExecWB OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, 140&
    Case 150
        wbExplorer.ExecWB OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, 150&
    Case 160
        wbExplorer.ExecWB OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, 160&
    Case 170
        wbExplorer.ExecWB OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, 170&
    Case 180
        wbExplorer.ExecWB OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, 180&
    Case 190
        wbExplorer.ExecWB OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, 190&
    Case 200
        wbExplorer.ExecWB OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, 200&
    End Select
    
    
End Sub


