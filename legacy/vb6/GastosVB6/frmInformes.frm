VERSION 5.00
Begin VB.Form frmInformes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes por pantalla"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   5985
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   240
      Picture         =   "frmInformes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton cmdInformeRegion 
      Height          =   495
      Left            =   3240
      Picture         =   "frmInformes.frx":0702
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton cmdInformeLitros 
      Height          =   495
      Left            =   3240
      Picture         =   "frmInformes.frx":0E04
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton cmdInformeAgrupado 
      Height          =   495
      Left            =   3240
      Picture         =   "frmInformes.frx":1506
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton btInformeConcepto 
      Height          =   495
      Left            =   3240
      Picture         =   "frmInformes.frx":1C08
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton btInformeMensual 
      Height          =   495
      Left            =   225
      Picture         =   "frmInformes.frx":230A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton btInformeGastos 
      Height          =   495
      Left            =   240
      Picture         =   "frmInformes.frx":2A0C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton btInformeAnual 
      Height          =   495
      Left            =   240
      Picture         =   "frmInformes.frx":310E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton btInformeDetallado 
      Height          =   495
      Left            =   240
      Picture         =   "frmInformes.frx":3810
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton btCerrar 
      BackColor       =   &H8000000B&
      Caption         =   "C&errar"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label lblInformeDescuentos 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Informe de Descuentos"
      Height          =   255
      Left            =   855
      TabIndex        =   18
      Top             =   2760
      Width           =   1800
   End
   Begin VB.Label lblInformeRegion 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Informe Región"
      Height          =   255
      Left            =   3840
      TabIndex        =   17
      Top             =   2160
      Width           =   1800
   End
   Begin VB.Label lblInformeLitros 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Informe Litros"
      Height          =   255
      Left            =   3840
      TabIndex        =   15
      Top             =   1560
      Width           =   1800
   End
   Begin VB.Label lblInformeAgrupado 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Informe Agrupado"
      Height          =   255
      Left            =   3840
      TabIndex        =   13
      Top             =   960
      Width           =   1800
   End
   Begin VB.Label lblInformeConcepto 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Informe  por  Concepto"
      Height          =   255
      Left            =   3840
      TabIndex        =   12
      Top             =   360
      Width           =   1800
   End
   Begin VB.Label etInformeMensual 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Informe Mensual"
      Height          =   255
      Left            =   840
      TabIndex        =   11
      Top             =   2160
      Width           =   1800
   End
   Begin VB.Label etInformeGastos 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Informe Gastos"
      Height          =   255
      Left            =   840
      TabIndex        =   10
      Top             =   1560
      Width           =   1800
   End
   Begin VB.Label etInformeAnual 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Informe Anual"
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   960
      Width           =   1800
   End
   Begin VB.Label etInformeDetallado 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Informe Detallado"
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   360
      Width           =   1800
   End
End
Attribute VB_Name = "frmInformes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btCerrar_Click()
    Unload Me
End Sub

Private Sub btInformeAnual_Click()
    frmInformeAnual.Show vbModal
End Sub

Private Sub btInformeConcepto_Click()
    frmInformeConcepto.Show vbModal
End Sub

Private Sub btInformeDetallado_Click()
    frmInfDetallado.Show vbModal
End Sub

Private Sub btInformeGastos_Click()
    frmInformeGastos.Show vbModal
End Sub

Private Sub btInformeMensual_Click()
    frmInformeMensual.Show vbModal
End Sub

Private Sub cmdInformeAgrupado_Click()
    frmInformeAgrupado.Show vbModal
End Sub

Private Sub cmdInformeLitros_Click()
    frmInformeLitrosGasolina.Show vbModal
End Sub

Private Sub cmdInformeRegion_Click()
    frmInformeRegion.Show vbModal
End Sub

Private Sub Command1_Click()
    frmInfDescuentos.Show vbModal
End Sub
