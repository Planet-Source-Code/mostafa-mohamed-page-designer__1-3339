Attribute VB_Name = "Module1"
'***************************************
'** Note converting page to html code is not
'my code but I it is already sent to planet source code
'by someone that i don't no his name,sorry because I can't
'credit him.
'***************************************
Public MyAgent As Object
Public imagemax As Integer
Public addctrl As String
Public bgsound As String
Public zoo As Integer
Public textmax As Integer
Public cmax As Integer
Public ctrltype As String
Public dirty As Boolean
Public shapemax As Integer
Public linemax As Integer
Public changes As Boolean
Public curfile As String
Public tcde(100) As String
Public icde(100) As String
Public runtime As Boolean
Public ccindexx As Integer
Public indexctrl As Integer
Public bgclr As OLE_COLOR
Public lclr As OLE_COLOR
Public vclr As OLE_COLOR
Public tclr As OLE_COLOR
Dim string1, string2
Private localTable As table
Private regionGroup() As region
Private tableArray() As Double
Private objectCounter As Integer
Private Const xLevel As Integer = 0, yLevel As Integer = 1, objectLevel As Integer = 2, drawnLevel As Integer = 3


Private Type table
    Width As Integer
    Height As Integer
    cellsWide As Integer
    cellsTall As Integer
    bgcolor As String
    html As String
    cellPadding As Integer
    End Type


Private Type region
    html As String
    Left As Long
    Top As Long
    Width As Long
    Height As Long
    bgcolor As String
    rowSpan As Long
    colSpan As Long
    End Type
    '***************************************************
    'end declarations
    'start code
    '***************************************************

Public Declare Function RasEnumConnections Lib "RasApi32.dll" Alias "RasEnumConnectionsA" (lpRasCon As Any, lpcb As Long, lpcConnections As Long) As Long


Public Declare Function RasGetConnectStatus Lib "RasApi32.dll" Alias "RasGetConnectStatusA" (ByVal hRasCon As Long, lpStatus As Any) As Long
    '
    Public Const RAS95_MaxEntryName = 256
    Public Const RAS95_MaxDeviceType = 16
    Public Const RAS95_MaxDeviceName = 32
    '


Public Type RASCONN95
    dwSize As Long
    hRasCon As Long
    szEntryName(RAS95_MaxEntryName) As Byte
    szDeviceType(RAS95_MaxDeviceType) As Byte
    szDeviceName(RAS95_MaxDeviceName) As Byte
    End Type
    '


Public Type RASCONNSTATUS95
    dwSize As Long
    RasConnState As Long
    dwError As Long
    szDeviceType(RAS95_MaxDeviceType) As Byte
    szDeviceName(RAS95_MaxDeviceName) As Byte
    End Type

Public Function Render() As String

    PrepareCrap 'get the crap ready For this new situation.
    MakeTable
    ClearAllRegions
    Render = localTable.html 'return the resulting html
End Function



Public Function AddRegion(html As String, Left As Double, Top As Double, Width As Double, Height As Double, bgcolor As String)

    objectCounter = objectCounter + 1
    ReDim Preserve regionGroup(objectCounter)
    If html = "" Then html = " "
    regionGroup(objectCounter - 1).html = html
    regionGroup(objectCounter - 1).Left = Left
    regionGroup(objectCounter - 1).Top = Top
    regionGroup(objectCounter - 1).Width = Width
    regionGroup(objectCounter - 1).Height = Height
    If bgcolor <> "" Then
    regionGroup(objectCounter - 1).bgcolor = bgcolor
End If
End Function



Private Function ClearAllRegions()

    objectCounter = 0
    Erase regionGroup()
End Function




Private Sub PrepareCrap()

    Erase tableArray()
    localTable.cellsWide = calculateCellsWide
    localTable.cellsTall = calculateCellsTall
    localTable.html = "" 'set html to nothing so that old rendering doesn't show up here
    localTable.Width = 0
    localTable.Height = 0
    ReDim tableArray(localTable.cellsWide, localTable.cellsTall, 4) 'resize the tablearray table


    For i = 0 To localTable.cellsWide


        For j = 0 To localTable.cellsTall
            tableArray(i, j, objectLevel) = -1
        Next j

    Next i

End Sub



Private Sub SortX()

    Dim edgeCoordinate As Integer


    For i = 0 To localTable.cellsWide - 1
        edgeCoordinate = 9999


        For j = 0 To (objectCounter - 1)


            If (i = 0) Then


                If (regionGroup(j).Left < edgeCoordinate) Then
                    edgeCoordinate = regionGroup(j).Left
                End If

            ElseIf ((regionGroup(j).Left < edgeCoordinate) And (regionGroup(j).Left > tableArray((i - 1), 0, xLevel))) Then
                edgeCoordinate = regionGroup(j).Left
            End If

        Next j



        If (edgeCoordinate <> 9999) Then
            If i = localTable.cellsWide Then Beep
            tableArray(i, 0, xLevel) = edgeCoordinate
        End If

    Next i

End Sub



Private Sub SortY()

    Dim edgeCoordinate As Integer


    For i = 0 To localTable.cellsTall - 1
        edgeCoordinate = 9999


        For j = 0 To (objectCounter - 1)


            If (i = 0) Then


                If (regionGroup(j).Top < edgeCoordinate) Then
                    edgeCoordinate = regionGroup(j).Top
                End If

            ElseIf ((regionGroup(j).Top < edgeCoordinate) And (regionGroup(j).Top > tableArray(0, (i - 1), yLevel))) Then
                edgeCoordinate = regionGroup(j).Top
            End If

        Next j



        If (edgeCoordinate <> 9999) Then
            tableArray(0, i, yLevel) = edgeCoordinate
        End If

    Next i

End Sub



Private Function LayoutTable()

    localTable.html = localTable.html & "<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=" & localTable.cellPadding & " BGCOLOR=" & localTable.bgcolor & ">"


    If tableArray(0, 0, yLevel) <> 0 Then 'only Do this if there is some height to give to the vertical offset
        localTable.html = localTable.html & "<TR>" 'start the row


        If (tableArray(0, 0, xLevel) <> 0) Then 'only print this first cell if there is some horizontal offset
            localTable.html = localTable.html & "<TD>" & vbCrLf
            localTable.html = localTable.html & "<IMG SRC=trans.gif HEIGHT=" & Chr(34) & tableArray(0, 0, yLevel) & Chr(34) & " WIDTH=" & Chr(34) & tableArray(0, 0, xLevel) & Chr(34) & ">"
            localTable.html = localTable.html & "</TD>"
        End If

        'for loop starts here,
        'this needs to go through and make cells and clearGifs with the g
        '     eneric height, and variable widths


        For j = 0 To localTable.cellsWide - 1
            localTable.html = localTable.html & "<TD>"
            localTable.html = localTable.html & "<IMG SRC=trans.gif HEIGHT=" & Chr(34) & "1" & Chr(34)


            If j < localTable.cellsWide - 1 Then
                localTable.html = localTable.html & " WIDTH=" & Chr(34) & (tableArray(j + 1, 0, xLevel) - tableArray(j, 0, xLevel)) & Chr(34) & ">"
            Else
                localTable.html = localTable.html & " WIDTH=" & Chr(34) & (localTable.Width - tableArray(j, 0, xLevel)) & Chr(34) & ">"
            End If

            localTable.html = localTable.html & "</TD>"
        Next j

        localTable.html = localTable.html & "</TR>"
    End If



    For i = 0 To localTable.cellsTall - 1
        localTable.html = localTable.html & "<TR>"


        For j = 0 To localTable.cellsWide - 1


            If ((tableArray(j, 0, xLevel) <> 0) And (j = 0)) Then 'only Do this is there is a horizontal width in the very first cell of the whole table
                localTable.html = localTable.html & "<TD>"
             
                localTable.html = localTable.html & "<IMG SRC=trans.gif WIDTH=" & Chr(34) & "1" & Chr(34) 'print that width


                If i < localTable.cellsTall - 1 Then
                    'here it is
                    localTable.html = localTable.html & " HEIGHT=" & Chr(34) & Abs(tableArray(0, i + 1, yLevel) - tableArray(0, i, yLevel)) & Chr(34) & ">"
                Else
                    localTable.html = localTable.html & " HEIGHT=" & Chr(34) & Abs(localTable.Height - tableArray(0, i, yLevel)) & Chr(34) & ">"
                End If

                localTable.html = localTable.html & "</TD>"
            End If



            If tableArray(j, i, objectLevel) <> -1 Then
                localTable.html = localTable.html & "<TD VALIGN=TOP "


                If regionGroup((tableArray(j, i, objectLevel))).colSpan > 1 Then
                    localTable.html = localTable.html & "COLSPAN=" & regionGroup(tableArray(j, i, objectLevel)).colSpan & " "
                End If



                If regionGroup(tableArray(j, i, objectLevel)).rowSpan > 1 Then
                    localTable.html = localTable.html & "ROWSPAN=" & regionGroup(tableArray(j, i, objectLevel)).rowSpan
                End If

                localTable.html = localTable.html & "><table valign=top align=left border=0 cellspacing=0 cellpadding=0 width=" & regionGroup(tableArray(j, i, objectLevel)).Width & " height=" & regionGroup(tableArray(j, i, objectLevel)).Height & " "


                If regionGroup(tableArray(j, i, objectLevel)).bgcolor <> "" Then
                    localTable.html = localTable.html & "bgcolor=" & regionGroup(tableArray(j, i, objectLevel)).bgcolor
                End If

                localTable.html = localTable.html & "><tr><td>"
                'here is where the actual object placement occurs
                localTable.html = localTable.html & regionGroup(tableArray(j, i, objectLevel)).html
                'here is where the actual object placement occurs
                localTable.html = localTable.html & "</td></tr></table></TD>"
            ElseIf ((tableArray(j, i, drawnLevel) <> 1) And (tableArray(j, i, objectLevel) = -1)) Then
                localTable.html = localTable.html & "<TD>"
                localTable.html = localTable.html & "</TD>"
            End If

        Next j

        localTable.html = localTable.html & "</TR>"
    Next i

    localTable.html = localTable.html & "</table>"

End Function



Private Sub FindTableDimensions()



    For i = 0 To objectCounter - 1


        If ((regionGroup(i).Left + regionGroup(i).Width) > localTable.Width) Then
            localTable.Width = regionGroup(i).Left + regionGroup(i).Width
        End If



        If ((regionGroup(i).Top + regionGroup(i).Height) > localTable.Height) Then
            localTable.Height = regionGroup(i).Top + regionGroup(i).Height
        End If

    Next i

    'everything looks good this far
End Sub



Private Function MakeTable()

    FindTableDimensions
    SortX 'put the sides (in one dimension) in order smallest to largest
    SortY
    'I think this last step went right, but maybe not
    'the problem seems to only show when using objects positioned at
    '     a 0 on an axis
    assignObjects 'mark objects as being in certain cells, and their spans
    LayoutTable
End Function



Private Sub assignObjects()



    For Y = 0 To localTable.cellsWide - 1


        For j = 0 To localTable.cellsTall - 1
            Dim k As Integer
            k = 0


            For k = 0 To objectCounter - 1


                If ((tableArray(Y, 0, xLevel) = regionGroup(k).Left) And (tableArray(0, j, yLevel) = regionGroup(k).Top)) Then
                    tableArray(Y, j, objectLevel) = k


                    doIt Int(Y), Int(j), Int(k)
                    End If

                Next k

            Next j

        Next Y

    End Sub



Private Sub doIt(cellXpos As Integer, cellYpos As Integer, objectNum As Integer)

    Dim rightNum As Integer, bottomNum As Integer
    rightNum = (regionGroup(objectNum).Left + regionGroup(objectNum).Width)
    bottomNum = (regionGroup(objectNum).Top + regionGroup(objectNum).Height)


    For i = cellXpos To localTable.cellsWide - 1


        If (tableArray(i, 0, xLevel) < rightNum) Then
            regionGroup(objectNum).colSpan = regionGroup(objectNum).colSpan + 1
        End If



        For j = cellYpos To localTable.cellsTall - 1


            If i = cellXpos Then


                If (tableArray(0, j, yLevel) < bottomNum) Then
                    regionGroup(objectNum).rowSpan = regionGroup(objectNum).rowSpan + 1
                End If

            End If



            If ((tableArray(0, j, yLevel) < bottomNum) And (tableArray(i, 0, xLevel) < rightNum)) Then
                tableArray(i, j, drawnLevel) = 1
            End If

        Next j

    Next i

End Sub



Private Function calculateCellsWide() As Integer

    Dim duplicateEdges As Integer


    For i = 0 To (objectCounter - 1)


        For j = i To (objectCounter - 1)


            If ((regionGroup(i).Left = regionGroup(j).Left) And (i <> j)) Then
                duplicateEdges = duplicateEdges + 1
                j = objectCounter + 4
            End If

        Next j



        If (regionGroup(i).Left = 0) Then
            duplicateEdges = duplicateEdges - 1
        End If

    Next i

    calculateCellsWide = (objectCounter - duplicateEdges)
End Function



Private Function calculateCellsTall() As Integer

    Dim duplicateEdges As Integer


    For i = 0 To (objectCounter - 1)


        For j = i To (objectCounter - 1)


            If ((regionGroup(i).Top = regionGroup(j).Top) And (i <> j)) Then
                duplicateEdges = duplicateEdges + 1
                j = objectCounter + 4
            End If

        Next j



        If (regionGroup(i).Top = 0) Then
            duplicateEdges = duplicateEdges - 1
        End If

    Next i

    calculateCellsTall = (objectCounter - duplicateEdges)
End Function
Sub savehtml(filename As String)
Dim i As Integer
Dim html As String
Dim no As Integer
Dim al As String
Dim clr As String
Dim btu As String
Dim ebtu As String
Dim str1 As String
Dim str3 As String
Dim str4 As String
Dim hcode As String
Dim str2 As String
no = FreeFile
Form1.ScaleMode = 3
ClearAllRegions

Open filename For Output As #no

For i = 0 To Form1.Controls.Count - 1
btu = ""
ebtu = ""
If TypeOf Form1.Controls(i) Is Image Then
If Form1.Controls(i).Index <> 0 Then

str2 = Form1.Controls(i).Tag
If str2 = "" Then
hcode = "<img src = '" & Form1.Controls(i).ToolTipText & "' width='" & Form1.Controls(i).Width & "' height='" & Form1.Controls(i).Height & "'>"
Else
hcode = "<a href='" & str2 & "'><img src = '" & Form1.Controls(i).ToolTipText & "' width='" & Form1.Controls(i).Width & "' height='" & Form1.Controls(i).Height & "'></a>"
End If
AddRegion hcode, Form1.Controls(i).Left, Form1.Controls(i).Top, Form1.Controls(i).Width, Form1.Controls(i).Height, ""
End If
End If
If TypeOf Form1.Controls(i) Is Label Then

If Form1.Controls(i).Index <> 0 Then
If Form1.Controls(i).ToolTipText = "Text" Then
Select Case Form1.Controls(i).Alignment
Case 0: al = "left"
Case 1: al = "right"
Case 2: al = "center"
End Select
If Form1.Controls(i).FontBold = True Then
btu = btu + "<strong>"
ebtu = ebtu + "</strong>"
End If
If Form1.Controls(i).FontItalic = True Then
btu = btu + "<em>"
ebtu = ebtu + "</em>"
End If
If Form1.Controls(i).FontUnderline = True Then
btu = btu + "<u>"
ebtu = ebtu + "</u>"
End If
clr = gethcolor(Form1.Controls(i).ForeColor)
If Form1.Controls(i).Tag = "" Then
hcode = "<p align=""" & al & """><font color=""" & clr & """ size=""" & Form1.Controls(i).FontSize / 3 & """face=""" & Form1.Controls(i).FontName & """>" & btu & Form1.Controls(i).Caption & ebtu & "</font></p>"
Else
hcode = "<a href='" & Form1.Controls(i).Tag & "'><p align=""" & al & """><font color=""" & clr & """ size=""" & Form1.Controls(i).FontSize / 3 & """face=""" & Form1.Controls(i).FontName & """>" & btu & Form1.Controls(i).Caption & ebtu & "</font></p></a>"
End If
If Form1.Controls(i).BackStyle = 0 Then
AddRegion hcode, Form1.Controls(i).Left, Form1.Controls(i).Top, Form1.Controls(i).Width, Form1.Controls(i).Height, ""
Else
clr = gethcolor(Form1.Controls(i).BackColor)
AddRegion hcode, Form1.Controls(i).Left, Form1.Controls(i).Top, Form1.Controls(i).Width, Form1.Controls(i).Height, clr
End If
End If
If Form1.Controls(i).ToolTipText = "Line" Then
clr = gethcolor(Form1.Controls(i).BackColor)
hcode = "<hr size='" & Form1.Controls(i).Height & "'color='" & clr & "'>"
AddRegion hcode, Form1.Controls(i).Left, Form1.Controls(i).Top, Form1.Controls(i).Width, Form1.Controls(i).Height, ""
End If
If Form1.Controls(i).ToolTipText = "Html code" Then
hcode = Form1.Controls(i).Tag
AddRegion hcode, Form1.Controls(i).Left, Form1.Controls(i).Top, Form1.Controls(i).Width, Form1.Controls(i).Height, ""
End If
End If
End If

Next i
str1 = gethcolor(bgclr)
str2 = gethcolor(vclr)
str3 = gethcolor(tclr)
str4 = gethcolor(lclr)
Print #no, "<body background='" & MDI.Picture11.Tag & "' bgcolor='" & str1 & "' vlink='" & str2 & "' text='" & str3 & "' link='" & str4 & "'></body>"
Print #no, "<bgsound src='" & bgsound & "' loop='infinite'>"
Print #no, "<title>" & Form1.Tag & "</title>"
Print #no, Render
Close #no
Form1.ScaleMode = 1
End Sub


Sub Savepage(filename As String)
'On Error Resume Next
Dim no As Integer
Dim i As Integer
no = FreeFile
'Unload Form2

Open filename For Output As #no
Print #no, "Page setup"
Write #no, Form1.Width
Write #no, Form1.Height
Write #no, MDI.Picture11.Tag
Write #no, Form1.Tag
Write #no, bgclr
Write #no, vclr
Write #no, tclr
Write #no, "none"
Write #no, bgsound
Write #no, lclr
Write #no, "none"
Write #no, "none"
Write #no, "none"
Write #no, "none"
Write #no, "none"
Write #no, "none"
Write #no, "none"
For i = 0 To Form1.Controls.Count - 1
If Not TypeOf Form1.Controls(i) Is PictureBox Then
If TypeOf Form1.Controls(i) Is Label And Form1.Controls(i).Index <> 0 Then
If Form1.Controls(i).ToolTipText = "Text" Then
Write #no, "Label"
Write #no, Form1.Controls(i).Caption
Write #no, Form1.Controls(i).Width
Write #no, Form1.Controls(i).Height
Write #no, Form1.Controls(i).Left
Write #no, Form1.Controls(i).Top
Write #no, Form1.Controls(i).BackColor
Write #no, Form1.Controls(i).ForeColor
Write #no, Form1.Controls(i).BackStyle
Write #no, Form1.Controls(i).FontName
Write #no, Form1.Controls(i).FontSize
Write #no, Form1.Controls(i).FontBold
Write #no, Form1.Controls(i).FontUnderline
Write #no, Form1.Controls(i).FontItalic
Write #no, Form1.Controls(i).Alignment
Write #no, tcde(Form1.Controls(i).Index)
Write #no, Form1.Controls(i).Tag
Write #no, Form1.Controls(i).ToolTipText
End If
If Form1.Controls(i).ToolTipText = "Html code" Then
Write #no, "Label"
Write #no, Form1.Controls(i).Caption
Write #no, Form1.Controls(i).Width
Write #no, Form1.Controls(i).Height
Write #no, Form1.Controls(i).Left
Write #no, Form1.Controls(i).Top
Write #no, Form1.Controls(i).BackColor
Write #no, Form1.Controls(i).ForeColor
Write #no, Form1.Controls(i).BackStyle
Write #no, Form1.Controls(i).FontName
Write #no, Form1.Controls(i).FontSize
Write #no, Form1.Controls(i).FontBold
Write #no, Form1.Controls(i).FontUnderline
Write #no, Form1.Controls(i).FontItalic
Write #no, Form1.Controls(i).Alignment
Write #no, tcde(Form1.Controls(i).Index)
Write #no, Form1.Controls(i).Tag
Write #no, Form1.Controls(i).ToolTipText
End If
If Form1.Controls(i).ToolTipText = "Line" Then
Write #no, "Label"
Write #no, Form1.Controls(i).Caption
Write #no, Form1.Controls(i).Width
Write #no, Form1.Controls(i).Height
Write #no, Form1.Controls(i).Left
Write #no, Form1.Controls(i).Top
Write #no, Form1.Controls(i).BackColor
Write #no, Form1.Controls(i).ForeColor
Write #no, Form1.Controls(i).BackStyle
Write #no, Form1.Controls(i).FontName
Write #no, Form1.Controls(i).FontSize
Write #no, Form1.Controls(i).FontBold
Write #no, Form1.Controls(i).FontUnderline
Write #no, Form1.Controls(i).FontItalic
Write #no, Form1.Controls(i).Alignment
Write #no, tcde(Form1.Controls(i).Index)
Write #no, Form1.Controls(i).Tag
Write #no, Form1.Controls(i).ToolTipText
End If
End If
If TypeOf Form1.Controls(i) Is Shape And Form1.Controls(i).Index <> 0 Then
Write #no, "Shape"
Write #no, Form1.Controls(i).Shape
Write #no, Form1.Controls(i).Width
Write #no, Form1.Controls(i).Height
Write #no, Form1.Controls(i).Left
Write #no, Form1.Controls(i).Top
Write #no, Form1.Controls(i).BackColor
Write #no, Form1.Controls(i).BorderColor
Write #no, Form1.Controls(i).BackStyle
Write #no, Form1.Controls(i).BorderStyle
Write #no, Form1.Controls(i).BorderWidth
Write #no, "none"
Write #no, "none"
Write #no, "none"
Write #no, "none"
Write #no, "none"
Write #no, "none"
Write #no, Form1.Controls(i).Tag
End If
If TypeOf Form1.Controls(i) Is Image And Form1.Controls(i).Index <> 0 Then
Write #no, "Image"
Write #no, Form1.Controls(i).Width
Write #no, Form1.Controls(i).Height
Write #no, Form1.Controls(i).Left
Write #no, Form1.Controls(i).Top
Write #no, Form1.Controls(i).ToolTipText
Write #no, "none"
Write #no, "none"
Write #no, "none"
Write #no, "none"
Write #no, "none"
Write #no, "none"
Write #no, "none"
Write #no, "none"
Write #no, "none"
Write #no, "none"
Write #no, icde(Form1.Controls(i).Index)
Write #no, Form1.Controls(i).Tag
End If
End If
Next i
Close #no

End Sub

Sub Openpage(filename As String)
On Error Resume Next
Unload Form2
Form2.Show
Form2.Hide
imagemax = 0
textmax = 0
cmax = 0
shapemax = 0
linemax = 0
Dim no As Integer
Dim i As Integer

no = FreeFile
Open filename For Input As #no
Do
Input #no, a
Input #no, b
Input #no, c
Input #no, d
Input #no, e
Input #no, f
Input #no, g
Input #no, h
Input #no, i
Input #no, j
Input #no, k
Input #no, l
Input #no, m
Input #no, n
Input #no, o
Input #no, p
Input #no, Q
Input #no, r
If a = "Page setup" Then
Form1.Width = b
Form1.Height = c
MDI.Picture11.Tag = d
MDI.Picture11.Picture = LoadPicture(d)
Form1.Tag = e
bgclr = f
Form1.BackColor = bgclr
vclr = g
tclr = h
bgsound = j
lclr = k
End If
If a = "Label" Then
If r = "Text" Then
textmax = textmax + 1
Load Form1.Label1(textmax)
Form1.Label1(textmax).Caption = b
Form1.Label1(textmax).Width = c
Form1.Label1(textmax).Height = d
Form1.Label1(textmax).Left = e
Form1.Label1(textmax).Top = f
Form1.Label1(textmax).BackStyle = i
Form1.Label1(textmax).BackColor = g
Form1.Label1(textmax).ForeColor = h
Form1.Label1(textmax).FontName = j
Form1.Label1(textmax).FontSize = k
Form1.Label1(textmax).FontBold = l
Form1.Label1(textmax).FontUnderline = m
Form1.Label1(textmax).FontItalic = n
Form1.Label1(textmax).Alignment = o
Form1.Label1(textmax).Tag = Q
Form1.Label1(textmax).ToolTipText = r
Form1.Label1(textmax).ZOrder 0
tcde(textmax) = p
Form1.Label1(textmax).Visible = True
textmax = textmax + 1
End If
If r = "Line" Then
linemax = linemax + 1
Load Form1.Label2(linemax)
Form1.Label2(linemax).Caption = b
Form1.Label2(linemax).Width = c
Form1.Label2(linemax).Height = d
Form1.Label2(linemax).Left = e
Form1.Label2(linemax).Top = f
Form1.Label2(linemax).BackStyle = i
Form1.Label2(linemax).BackColor = g
Form1.Label2(linemax).ForeColor = h
Form1.Label2(linemax).FontName = j
Form1.Label2(linemax).FontSize = k
Form1.Label2(linemax).FontBold = l
Form1.Label2(linemax).FontUnderline = m
Form1.Label2(linemax).FontItalic = n
Form1.Label2(linemax).Alignment = o
Form1.Label2(linemax).ToolTipText = r
Form1.Label2(linemax).Tag = Q
Form1.Label2(linemax).ZOrder 0
Form1.Label2(linemax).Visible = True
linemax = linemax + 1
End If
If r = "Html code" Then
cmax = cmax + 1
Load Form1.Label3(cmax)
Form1.Label3(cmax).Caption = b
Form1.Label3(cmax).Width = c
Form1.Label3(cmax).Height = d
Form1.Label3(cmax).Left = e
Form1.Label3(cmax).Top = f
Form1.Label3(cmax).BackStyle = i
Form1.Label3(cmax).BackColor = g
Form1.Label3(cmax).ForeColor = h
Form1.Label3(cmax).FontName = j
Form1.Label3(cmax).FontSize = k
Form1.Label3(cmax).FontBold = l
Form1.Label3(cmax).FontUnderline = m
Form1.Label3(cmax).FontItalic = n
Form1.Label3(cmax).Alignment = o
Form1.Label3(cmax).ToolTipText = r
Form1.Label3(cmax).Tag = Q
Form1.Label3(cmax).ZOrder 0
Form1.Label3(cmax).Visible = True
cmax = cmax + 1
End If
End If

If a = "Shape" Then
shapemax = shapemax + 1
Load Form1.Shape1(shapemax)
Form1.Shape1(shapemax).Shape = b
Form1.Shape1(shapemax).Width = c
Form1.Shape1(shapemax).Height = d
Form1.Shape1(shapemax).Left = e
Form1.Shape1(shapemax).Top = f
Form1.Shape1(shapemax).BackStyle = i
Form1.Shape1(shapemax).BackColor = g
Form1.Shape1(shapemax).BorderColor = h
Form1.Shape1(shapemax).BorderStyle = j
Form1.Shape1(shapemax).BorderWidth = k
Form1.Shape1(shapemax).Visible = True
shapemax = shapemax + 1
End If

If a = "Image" Then
imagemax = imagemax + 1
Load Form1.Image1(imagemax)
Form1.Image1(imagemax).Width = b
Form1.Image1(imagemax).Height = c
Form1.Image1(imagemax).Left = d
Form1.Image1(imagemax).Top = e
Form1.Image1(imagemax).ToolTipText = f
Form1.Image1(imagemax).Stretch = True
Form1.Image1(imagemax).Visible = True
Form1.Image1(imagemax).Picture = LoadPicture(f)
Form1.Image1(imagemax).Tag = r
Form1.Image1(imagemax).ZOrder 0
icde(imagemax) = p
imagemax = imagemax + 1
End If


Loop Until EOF(no)
Close #no
End Sub
Sub dec(filename As String)

Dim Message As String
Dim TryPass As String
Dim charnum As Currency, randominteger As Currency
Dim singlechar As String * 1
Dim keyvalue As Currency
Dim secondkey As Currency
Dim CurrChar As String
Dim msg As String
Dim ctxt As Integer
Dim filenum As Currency
Dim X As Currency
Dim i As Currency
filenum = FreeFile
 filename$ = filename
Open filename$ For Binary As #filenum
    For i = 1 To LOF(filenum)
      Get #filenum, i, singlechar
      charnum = Asc(singlechar)
      randominteger = Int(256 * Rnd)
      charnum = charnum Xor randominteger
      singlechar = Chr$(charnum)
      Put #filenum, i, singlechar
    Next i
  Close #filenum
End Sub
Sub krypt(filename)
Dim Message As String
Dim TryPass As String
Dim charnum As Currency, randominteger As Currency
Dim singlechar As String * 1
Dim keyvalue As Currency
Dim secondkey As Currency
Dim CurrChar As String
Dim msg As String
Dim ctxt As Integer
Dim Q As Currency
Dim filenum As Currency
Dim X As Currency
Dim i As Currency
Dim xx As Currency












    filenum = FreeFile
    X = Rnd(-keyvalue)
    
 
    
   

    Open filename For Binary As #filenum     'open the file name for output/input.
    For i = 1 To LOF(filenum)
      Get #filenum, i, singlechar
      charnum = Asc(singlechar)
      randominteger = Int(256 * Rnd)
      charnum = charnum Xor randominteger
      singlechar = Chr$(charnum)
      Put #filenum, i, singlechar
    Next i
  Close #filenum
End Sub
Function getcommand(searchin As String) As String
On Error Resume Next
Dim n As Integer, s As String

n = InStr(1, searchin, ",", vbTextCompare)
s = Mid(searchin, 1, n - 1)
s = Format(s, ">")
getcommand = s
End Function

Function getval(searchin As String) As String
On Error Resume Next
Dim n As Integer, s As String
n = InStr(1, searchin, ",", vbTextCompare)
X = Len(searchin)
s = Mid(searchin, n + 1, X - 1)
getval = s
End Function

Function gethcolor(colorlong As Long) As String
On Error Resume Next
Dim Red As Long, Green As Long, Blue As Long

Red = colorlong And &HFF&
Green = (colorlong And &HFF00&) \ 256
Blue = (colorlong And &HFF0000) \ 65536
gethcolor = "#" + Hex(Red) + Hex(Green) + Hex(Blue)


    If gethcolor = "#000" Then
         gethcolor = "#000000"
    End If

End Function
