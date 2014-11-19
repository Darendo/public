VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2790
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   2790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkBackfillMode 
      Alignment       =   1  'Right Justify
      Caption         =   "Back-fill mode:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   155
      TabIndex        =   2
      Top             =   1140
      Width           =   2080
   End
   Begin VB.CommandButton cmdAction 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   360
      Index           =   1
      Left            =   1440
      TabIndex        =   4
      Top             =   1740
      Width           =   990
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   360
      Index           =   0
      Left            =   420
      TabIndex        =   3
      Top             =   1740
      Width           =   990
   End
   Begin VB.TextBox txtDateModifiedOffset 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Text            =   "10"
      ToolTipText     =   "Enter a numeric value in minutes"
      Top             =   660
      Width           =   495
   End
   Begin VB.TextBox txtFrequency 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Text            =   "3"
      ToolTipText     =   "Enter a numeric value in seconds"
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Modified Offset:"
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   6
      Top             =   720
      Width           =   1665
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Frequency:"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   5
      Top             =   300
      Width           =   885
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bCancel As Boolean

Private Sub Form_Load()
    m_bCancel = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    m_bCancel = (UnloadMode = vbFormControlMenu)
    Cancel = 1
    Hide
End Sub

Private Sub cmdAction_Click(Index As Integer)
    m_bCancel = (Index = 1)
    Hide
End Sub

Private Sub txtDateModifiedOffset_GotFocus()
    txtDateModifiedOffset.SelStart = 0
    txtDateModifiedOffset.SelLength = Len(txtDateModifiedOffset.Text)
End Sub

Private Sub txtDateModifiedOffset_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode < 48 Or KeyCode > 57 Then KeyCode = 0
End Sub

Private Sub txtFrequency_GotFocus()
    txtFrequency.SelStart = 0
    txtFrequency.SelLength = Len(txtFrequency.Text)
End Sub

Private Sub txtFrequency_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode < 48 Or KeyCode > 57 Then KeyCode = 0
End Sub


Public Property Get Cancel() As Boolean
    Cancel = m_bCancel
End Property

Public Property Get Frequency() As Long
Dim lTmp As Long
    lTmp = CLng(Val(txtFrequency.Text))
    If lTmp <= 0 Then lTmp = 1
    Frequency = lTmp
End Property

Public Property Let Frequency(ByVal Value As Long)
    If Value <= 0 Then Value = 1
    txtFrequency.Text = CStr(Value)
End Property

Public Property Get DateModifiedOffset() As Long
Dim lTmp As Long
    lTmp = CLng(Val(txtDateModifiedOffset.Text))
    If lTmp < 0 Then lTmp = 0
    DateModifiedOffset = lTmp
End Property

Public Property Let DateModifiedOffset(ByVal Value As Long)
    If Value < 0 Then Value = 0
    txtDateModifiedOffset.Text = CStr(Value)
End Property

Public Property Get BackfillMode() As Boolean
    BackfillMode = (chkBackfillMode.Value = vbChecked)
End Property

Public Property Let BackfillMode(Value As Boolean)
    chkBackfillMode.Value = IIf(Value, vbChecked, vbUnchecked)
End Property
