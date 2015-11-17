VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FechamentoCupomFrm 
   Caption         =   "Fechamento do Cupom Fiscal"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5280
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
   ScaleHeight     =   3690
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton btnDadosCheque 
      Caption         =   "&Dados do Cheque"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   5055
   End
   Begin VB.ComboBox cmbForma 
      Height          =   330
      ItemData        =   "FechamentoCupomFrm.frx":0000
      Left            =   480
      List            =   "FechamentoCupomFrm.frx":000D
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   2900
   End
   Begin MSDataGridLib.DataGrid dbgFormas 
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3625
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   "Forma de Pagamento"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   "Valor"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   2894,74
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton btnExcluir 
      Caption         =   "&Excluir"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton btnInserir 
      Caption         =   "&Inserir"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "FechamentoCupomFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rs As ADODB.Recordset
Public fechouOK As Boolean
Public numCartoes As Integer

Private Sub carregarRecordSet()
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Fields.Append "Forma", adBSTR, 20
    rs.Fields.Append "Valor", adDouble
    rs.Open
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        While Not rs.EOF
            rs.Delete
        Wend
    End If
    
    With dbgFormas
        Set .DataSource = Nothing
        Set .DataSource = rs
        .Refresh
    End With
End Sub

Private Sub btnCancelar_Click()
    fechouOK = False
    'Undload Me
End Sub

Private Sub btnExcluir_Click()
    rs.Delete adAffectCurrent
End Sub

Private Sub btnInserir_Click()
    rs.AddNew
End Sub

Private Sub btnOk_Click()
    fechouOK = True
End Sub

Private Sub cmbForma_Click()
    If cmbForma.Visible = True Then
        dbgFormas.Columns(0).Text = cmbForma.Text
        rs.Fields("Forma").Value = cmbForma.Text
    End If
End Sub

Private Sub dbgFormas_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If dbgFormas.Col = 0 Then
        cmbForma.Visible = True
        cmbForma.Top = dbgFormas.Columns(0).Top
        dbgFormas.RowHeight = cmbForma.Height
        If dbgFormas.Row > 0 Then
            cmbForma.Top = dbgFormas.Columns(0).Top + dbgFormas.Top + dbgFormas.RowTop(dbgFormas.Row)
            cmbForma.Top = cmbForma.Top - 200
        Else
            cmbForma.Top = 845
        End If
        cmbForma.Width = dbgFormas.Columns(0).Width
        cmbForma.Left = dbgFormas.Columns(0).Left + dbgFormas.Left
    Else
        cmbForma.Visible = False
    End If
End Sub

Private Sub Form_Load()
    fechouOK = False
    carregarRecordSet
End Sub
