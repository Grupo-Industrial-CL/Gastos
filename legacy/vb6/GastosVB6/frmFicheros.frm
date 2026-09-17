VERSION 5.00
Begin VB.Form frmFicheros 
   Caption         =   "Exportacion de datos a ficheros"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5700
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   5700
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btCerrar 
      BackColor       =   &H8000000B&
      Caption         =   "C&errar"
      CausesValidation=   0   'False
      Height          =   735
      Left            =   4800
      Picture         =   "frmFicheros.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3720
      Width           =   855
   End
End
Attribute VB_Name = "frmFicheros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btCerrar_Click()
Unload Me
End Sub
