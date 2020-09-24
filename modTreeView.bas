Attribute VB_Name = "modTreeView"
'***************************************************************************
' Please dont remove the comments if you use this in your own programs.    *
' It would benefit the whole comutity if all info was left intact in all   *
' the free examples on the net.  Thank you.                                *
'***************************************************************************

'This example builds on Srinis example found on www.vb-helper.com.
'Srinis example loaded data to a treeview from a text file with tabs
'denoting indentation.
'My contribution is the save treeview to file function.
'It only saves one level of childs, it was all I needed for my other project.
'Srinis example loads more than one level of childs. I am shure some small
'modifications to my example will enable it to save multiple levels of childs,
'but thats your jobb ;)
'If you make that modification, I will be pleased if you sends it to me at
'trond.sorensen@bi.no

'Now this is a compleat (allmost) set of functions working both ways.
'All credits should go to Srini.

Option Explicit

' Load a TreeView control from a file that uses tabs
' to show indentation.
' SRINIS WORLD
' haisrini@ email.com

Sub LoadTreeViewFromFile(ByVal file_name As String, ByVal trv As TreeView)
Dim fnum As Integer
Dim text_line As String
Dim level As Integer
Dim tree_nodes() As Node
Dim num_nodes As Integer

    fnum = FreeFile
    Open file_name For Input As fnum

    frmMain.TVProjects.Nodes.Clear
    Do While Not EOF(fnum)
        ' Get a line.
        Line Input #fnum, text_line

        ' Find the level of indentation.
        level = 1
        Do While Left$(text_line, 1) = vbTab
            level = level + 1
            text_line = Mid$(text_line, 2)
        Loop

        ' Make room for the new node.
        If level > num_nodes Then
            num_nodes = level
            ReDim Preserve tree_nodes(1 To num_nodes)
        End If

        ' Add the new node.
        If level = 1 Then 'parent
            Set tree_nodes(level) = trv.Nodes.Add(, , , text_line, "Folder")
        Else    'child
            Set tree_nodes(level) = trv.Nodes.Add(tree_nodes(level - 1), tvwChild, , text_line, "TextFile")
            tree_nodes(level).EnsureVisible
        End If
    Loop

    Close fnum

End Sub

' Save a TreeView control to a file that uses tabs
' to show indentation.
' Code by Trond Sørensen.
' trond.sorensen@bi.no

Sub SaveTreeViewToFile(ByVal file_name As String, ByVal trv As TreeView)
Dim fnum As Integer
'Dim text_line As String
'Dim level As Integer
'Dim tree_nodes() As Node
Dim num_nodes As Integer
        
    fnum = FreeFile
    Open file_name For Output As fnum
        For num_nodes = 1 To trv.Nodes.Count
            ' if the node uses a folder as icon it is a parent
            If trv.Nodes.Item(num_nodes).Image = "Folder" Then
                Print #fnum, trv.Nodes.Item(num_nodes)
            Else    ' if the node is not using a folder as icon it is a child
                Print #fnum, vbTab & trv.Nodes.Item(num_nodes)
            End If
        Next
    Close fnum

End Sub

Sub AddNewNode(ByVal trv As TreeView)
Dim Name As String
Dim Person As Node
Dim Group As Node

    Name = InputBox("Node Name", "New Node", "")
    If Name = "" Then Exit Sub
    
    ' Find the group that should hold the new node.
    ' if the selected item is a parent
    If trv.SelectedItem.Image = "Folder" Then
        Set Group = trv.SelectedItem
        Set Person = trv.Nodes.Add(Group, tvwChild, , Name, "TextFile")
    Else ' if selected item is a child
        Set Group = trv.SelectedItem.Parent
        Set Person = trv.Nodes.Add(Group, tvwChild, , Name, "TextFile")
    End If
    
    Person.EnsureVisible

End Sub

Public Sub AddNewGroup(ByVal trv As TreeView)
Dim Name As String
Dim Group As Node

    Name = InputBox("Node Name", "New Node", "")
    If Name = "" Then Exit Sub
    
    ' Add a new parent
        Set Group = trv.Nodes.Add(, , , Name, "Folder")
    
    Group.EnsureVisible

End Sub
