VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTreeView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   3510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDown 
      Caption         =   "Down"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdUP 
      Caption         =   "Up"
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   5640
      Width           =   1215
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   9551
      _Version        =   393217
      HideSelection   =   0   'False
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "frmTreeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDown_Click()
  MoveNode TreeView1, TreeView1.SelectedItem, "DOWN"
End Sub

Private Sub cmdUP_Click()
  MoveNode TreeView1, TreeView1.SelectedItem, "UP"
End Sub

Private Sub Form_Load()
Dim nodX As Node, nodA As Node
Dim i As Integer, j As Integer, k As Integer
  
  'Populate our TreeView with dummy nodes
  '3 levels deep so that we can play:
  For i = 1 To 10
    Set nodX = TreeView1.Nodes.Add(, , , "Grand Parent #" & i)
    For j = 1 To 5
      Set nodA = TreeView1.Nodes.Add(nodX.Index, tvwChild, , "Parent #" & j & " Child of #" & i)
      If j = 3 Then
        For k = 1 To 2
          TreeView1.Nodes.Add nodA.Index, tvwChild, , "GrandChild of Parent #" & j
        Next
      End If
    Next
  Next
End Sub

'This is our recursive function to find the children
'Of a node, and that' nodes children, and so on.
Private Sub GetChildren(tvw As TreeView, nodN As Node, nodP As Node)
Dim nodC As Node, nodT As Node
Dim i As Integer

  With tvw
    'For each children in the tree
    For i = 1 To nodN.Children
      'If it's the first child:
      If i = 1 Then
        'Add the node:
        Set nodC = .Nodes.Add(nodP.Index, tvwChild, , nodN.Child.Text)
        'Set us up for the next child:
        Set nodT = nodN.Child.Next
        'Get the added nodes children:
        If nodN.Child.Children <> 0 Then
          GetChildren tvw, nodN.Child, nodC
        End If
      'It's not the first child:
      Else
        On Error Resume Next
        'Add the node:
        Set nodC = .Nodes.Add(nodP.Index, tvwChild, , nodT.Text)
        'Get the added nodes children:
        If nodT.Children <> 0 Then
          GetChildren tvw, nodT, nodC
        End If
        'Set us up again:
        Set nodT = nodT.Next
      End If
    Next
  End With
End Sub

Private Sub MoveNode(tvw As TreeView, nodX As Node, Direction As String)
Dim nodN As Node
Dim strKey As String
  
  'All we do here is copy the node and set it as the previous
  'Nodes previous node. A little confusing, but it works.
  'We then add all the children and delete the original
  'Node
  
  With tvw
    Select Case Direction
      Case "UP"
        If Not nodX.Previous Is Nothing Then
          Set nodN = .Nodes.Add(nodX.Previous, tvwPrevious, , nodX.Text)
        Else
          Exit Sub
        End If
      Case "DOWN"
        If Not nodX.Next Is Nothing Then
          Set nodN = .Nodes.Add(nodX.Next, tvwNext, , nodX.Text)
        Else
          Exit Sub
        End If
    End Select
      
    nodN.Selected = True
      
    If nodX.Children <> 0 Then
      GetChildren tvw, nodX, nodN
    End If
      
    strKey = nodX.Key
    .Nodes.Remove nodX.Index
    Set nodX = Nothing
    nodN.Key = strKey
  End With
End Sub
