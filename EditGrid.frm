VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   7320
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEdit 
      Height          =   240
      Left            =   1620
      TabIndex        =   1
      Top             =   4215
      Visible         =   0   'False
      Width           =   1560
   End
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   4680
      Left            =   -15
      TabIndex        =   0
      Top             =   15
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   8255
      _Version        =   393216
      Rows            =   20
      Cols            =   8
      FixedRows       =   0
      FixedCols       =   0
      Appearance      =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'===============================================================================
'Purpose:       This example shows how to create the illusion of "in cell" eon
'               a plain old flexgrid. Why didn't MS do this?
'Returns:
'Created By:    Matthew M. Roberts
'Date:          5/25/2001
'Comments:      This could be made into a simple custom control pretty easily.
'===============================================================================

Private Sub Grd_SelChange()
   
   With txtEdit
        .Top = Grd.CellTop + Grd.Top
        .Left = Grd.CellLeft + Grd.Left
        .Height = Grd.CellHeight
        .Width = Grd.CellWidth
        .Text = Grd.Text
        .Visible = True
        .SetFocus
        'Lock grid because scrolling will mess up the textbox alignment
        Grd.Enabled = False
    End With
    
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    '   Pressing "Enter" ends the edit and re-enables the grid.
    If KeyAscii = 13 Then
        Grd.Enabled = True
        txtEdit.Visible = False
        With Grd
          .Text = txtEdit.Text
        End With
    End If

End Sub


