VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ListView 2 Word 2003 - by: Wouter"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   9255
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Options:"
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Width           =   9015
      Begin VB.CheckBox chkCloseDocument 
         Caption         =   "Close Word Document after clipboard move"
         Enabled         =   0   'False
         Height          =   255
         Left            =   4560
         TabIndex        =   7
         Top             =   600
         Width           =   3735
      End
      Begin VB.CheckBox chkMoveToClipboard 
         Caption         =   "Move the table to the clipboard"
         Height          =   255
         Left            =   4560
         TabIndex        =   6
         Top             =   240
         Width           =   4215
      End
      Begin VB.CheckBox chkBoldHeaders 
         Caption         =   "Use bold Column Headers"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   2295
      End
      Begin VB.CheckBox chkColumnHeaders 
         Caption         =   "Include the Column Headers"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Height          =   375
      Left            =   7680
      TabIndex        =   1
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "ListView:"
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9015
      Begin MSComctlLib.ListView lstViewMain 
         Height          =   3615
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   6376
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Column 1"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Column 2"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Column 3"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Column 4"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Column 5"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Column 6"
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    '
    ' ============================================================================
    ' Programmed by Wouter, wouter@vanhezel.nl
    ' Use every line of code for free! Want to add it in a commercial application?
    ' Just send me a email! Thanks!
    ' ============================================================================
    '

Private Sub chkColumnHeaders_Click()
    ' Check if the option is checked, if so Enable the sub options
    ' ============================================================================
    If chkColumnHeaders.Value = vbChecked Then
        chkBoldHeaders.Enabled = True
    Else
        chkBoldHeaders.Enabled = False
        chkBoldHeaders.Value = vbUnchecked
    End If
    
End Sub

Private Sub chkMoveToClipboard_Click()
    ' Check if the option is checked, if so Enable the sub options
    ' ============================================================================
    If chkMoveToClipboard.Value = vbChecked Then
        chkCloseDocument.Enabled = True
    Else
        chkCloseDocument.Enabled = False
        chkCloseDocument.Value = vbUnchecked
    End If
End Sub

Private Sub cmdExport_Click()
    ' Start Word and export the table to it
    ' ============================================================================
    Dim objWord As New Word.Application
    Dim intColumnCount As Integer              ' Column Count
    Dim intRowCount As Integer                 ' Row count
    Dim i As Integer, y As Integer             ' Just some counters
    
    intColumnCount = lstViewMain.ColumnHeaders.Count
    
    ' Inlcude a extra line if we want to have the Columnheaders included
    If chkColumnHeaders.Value = vbChecked Then
        intRowCount = lstViewMain.ListItems.Count + 1
    Else
        intRowCount = lstViewMain.ListItems.Count
    End If
    
    ' Add a new document to the existing Word
    objWord.Documents.Add , , wdNewBlankDocument
    objWord.Visible = True
    
    With objWord.Selection
        ' Add the table to the Word Document
        .Tables.Add .Range, intRowCount, intColumnCount, wdWord9TableBehavior, wdAutoFitFixed
        
        ' Include the Columnheaders
        If chkColumnHeaders.Value = vbChecked Then
            For i = 1 To intColumnCount
                If chkBoldHeaders.Value = vbChecked Then
                    .Font.Bold = wdToggle
                    .TypeText lstViewMain.ColumnHeaders.Item(i).Text
                    .Font.Bold = wdToggle
                Else
                    .TypeText lstViewMain.ColumnHeaders.Item(i).Text
                End If
                .MoveRight
            Next i
            ' The Rowcounter needs to be updated. Otherwise you will get extra line.
            ' You dont want to have that!
            intRowCount = intRowCount - 1
        End If
        
        ' i = the Row counter, every row needs to be added. Dûh.
        For i = 1 To intRowCount
            .TypeText lstViewMain.ListItems(i).Text
            .MoveRight
            
            ' y = the columncount
            For y = 1 To intColumnCount - 1
                .TypeText lstViewMain.ListItems(i).SubItems(y)
                .MoveRight
            Next y
        Next i
        
        ' Do we need to move the whole table to the clipboard?
        If chkMoveToClipboard.Value = vbChecked Then
            .MoveLeft wdCharacter, intColumnCount, wdExtend
            
            ' If the Column headers are included we need to copy one more line
            If chkColumnHeaders.Value = vbChecked Then
                .MoveUp wdLine, intRowCount + 1, wdExtend
            Else
                .MoveUp wdLine, intRowCount, wdExtend
            End If
            
            ' Move to clipboard
            .Cut
            
            ' Do we need to close the document?
            If chkCloseDocument.Value = vbChecked Then
                objWord.ActiveDocument.Close 0
            End If
        End If
        
    End With
End Sub

Private Sub Form_Load()
    ' Load bogus in the ListView
    ' ============================================================================
    Dim intKeyIndex As Integer      ' For the Key Index
    Dim i As Integer                ' Just a counter
    
    intKeyIndex = 1
    
    For i = 1 To 15                 ' Add 15 lines
        With lstViewMain.ListItems
            ' Fill the lstViewMain
            .Add intKeyIndex, "key" & intKeyIndex, "Line: " & intKeyIndex
            .Item(intKeyIndex).SubItems(1) = "Foo"
            .Item(intKeyIndex).SubItems(2) = "Bar"
            .Item(intKeyIndex).SubItems(3) = "Foo"
            .Item(intKeyIndex).SubItems(4) = "Bar"
            .Item(intKeyIndex).SubItems(5) = "Foo"
            
        End With
        
        ' Update counter
        intKeyIndex = intKeyIndex + 1
    Next i
    
End Sub
