VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmExplorer 
   BorderStyle     =   0  'None
   Caption         =   "Visor"
   ClientHeight    =   12360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15585
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12360
   ScaleWidth      =   15585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   2
      Text            =   "100"
      Top             =   735
      Width           =   1335
   End
   Begin SHDocVwCtl.WebBrowser wbExplorer 
      CausesValidation=   0   'False
      Height          =   10815
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1440
      Width           =   15375
      ExtentX         =   27120
      ExtentY         =   19076
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
   Begin VB.Shape shRecuadroForm 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   135
   End
   Begin VB.Image imgCursor 
      Height          =   480
      Left            =   5880
      Picture         =   "frmmExplorer.frx":0000
      Top             =   0
      Visible         =   0   'False
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
      Left            =   12120
      TabIndex        =   3
      Top             =   1200
      Width           =   3345
   End
   Begin VB.Image imgMenos 
      Height          =   480
      Left            =   600
      Picture         =   "frmmExplorer.frx":08CA
      ToolTipText     =   "Salir"
      Top             =   720
      Width           =   480
   End
   Begin VB.Image imgMas 
      Height          =   480
      Left            =   2880
      Picture         =   "frmmExplorer.frx":1194
      ToolTipText     =   "Salir"
      Top             =   720
      Width           =   480
   End
   Begin VB.Image imgSalir 
      Height          =   360
      Left            =   15000
      Picture         =   "frmmExplorer.frx":1A5E
      ToolTipText     =   "Salir"
      Top             =   120
      Width           =   360
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      BackColor       =   &H00E2C07A&
      Caption         =   "Visor de tickets"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Index           =   8
      Left            =   240
      TabIndex        =   1
      Top             =   75
      Width           =   2745
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H80000001&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E2C07A&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   15585
   End
End
Attribute VB_Name = "frmExplorer"
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

Private Sub Form_Load()
    
    If DisplayMonitorCount = 1 Then
        Me.Left = (Screen.Width / 2) - (Me.Width / 2)
        Me.Top = (Screen.Height / 2) - (Me.Height / 2)
    Else
        If MDIPrincipal.Left + 120 < Screen.Width Then
        'El MDI está en la pantalla 1, lo centro en la pantalla 1
            Me.Left = (Screen.Width / 2) - (Me.Width / 2)
            Me.Top = (Screen.Height / 2) - (Me.Height / 2)
        Else
            Me.Left = gAnchoPantallaPrincipal + (((gAnchoPantallaVirtual - gAnchoPantallaPrincipal) - Me.Width) \ 2)
            Me.Top = (Screen.Height - Me.Height) \ 2
        End If
    End If
    
    If msRuta = "" Then Exit Sub
    Select Case miAplicacion
    Case 0
        lblTexto(8).Caption = "Visor de tickets"
    Case 1
        lblTexto(8).Caption = "Visor de imágenes"
    Case Else
        lblTexto(8).Caption = "Visor"
    End Select
    shRecuadroForm.Width = Me.Width - 5
    shRecuadroForm.Height = Me.Height - 5
    shRecuadroForm.Left = 5
    shRecuadroForm.Top = 5
    wbExplorer.Navigate msRuta
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = DEFAULT
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

