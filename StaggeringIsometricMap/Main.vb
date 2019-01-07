Public Class Main

    Dim Drag As Boolean
    Dim Xaxys As Integer
    Dim Yaxys As Integer

    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        draw_map()
        enableDragControl()
        enableDragForm()
    End Sub

    Private Sub enableDragForm()
        AddHandler Me.MouseUp, Sub(e_sender, eventargs)
                                   disableDrag()
                               End Sub
        AddHandler Me.MouseDown, Sub(e_sender, eventargs)
                                     enableDrag()
                                 End Sub
        AddHandler Me.MouseMove, Sub(e_sender, eventargs)
                                     moveControl()
                                 End Sub
    End Sub

    Private Sub enableDragControl()
        For Each c As Control In Me.Controls
            AddHandler c.MouseUp, Sub(e_sender, eventargs)
                                      disableDrag()
                                  End Sub
            AddHandler c.MouseDown, Sub(e_sender, eventargs)
                                        enableDrag()
                                    End Sub
            AddHandler c.MouseMove, Sub(e_sender, eventargs)
                                        moveControl()
                                    End Sub
            If TypeOf (c) Is Panel Then
                For Each panelctl As Control In c.Controls
                    AddHandler panelctl.MouseUp, Sub(e_sender, eventargs)
                                                     disableDrag()
                                                 End Sub
                    AddHandler panelctl.MouseDown, Sub(e_sender, eventargs)
                                                       enableDrag()
                                                   End Sub
                    AddHandler panelctl.MouseMove, Sub(e_sender, eventargs)
                                                       moveControl()
                                                   End Sub
                Next
            End If
        Next
    End Sub

    Public Sub draw_map()
        Dim g As Graphics = IsometricGrid.CreateGraphics()

        Dim MapWidth As Integer = 15
        Dim MapHeight As Integer = 17

        Dim cellWidth As Double = 40
        Dim cellHeight As Double = Math.Ceiling(cellWidth / 2)

        Dim offsetX As Integer = Convert.ToInt32((Width - ((MapWidth + 0.5) * cellWidth)) / 2)
        Dim offsetY = Convert.ToInt32((Height - ((MapHeight + 0.5) * cellHeight)) / 2)

        Dim medianCellH As Double = cellHeight / 2
        Dim medianCellW As Double = cellWidth / 2

        Dim count As Integer = 0

        For y As Integer = 0 To 2 * MapHeight - 1

            If (y Mod 2) = 0 Then

                For x As Integer = 0 To MapWidth - 1
                    Dim left As Point = New Point(CInt((offsetX + x * cellWidth)), CInt((offsetY + y * medianCellH + medianCellH)))
                    Dim top As Point = New Point(CInt((offsetX + (x * cellWidth) + medianCellW)), CInt((offsetY + (y * medianCellH))))
                    Dim right As Point = New Point(CInt((offsetX + x * cellWidth + cellWidth)), CInt((offsetY + y * medianCellH + medianCellH)))
                    Dim down As Point = New Point(CInt((offsetX + (x * cellWidth) + medianCellW)), CInt((offsetY + (y * medianCellH) + cellHeight)))
                    renderCell(left, top, right, down, count, Color.DimGray)
                    count += 1
                Next
            Else

                For x As Integer = 0 To MapWidth - 2
                    Dim left As Point = New Point(CInt((offsetX + x * cellWidth + medianCellW)), CInt((offsetY + y * medianCellH + medianCellH)))
                    Dim top As Point = New Point(CInt((offsetX + x * cellWidth + cellWidth)), CInt((offsetY + y * medianCellH)))
                    Dim right As Point = New Point(CInt((offsetX + x * cellWidth + cellWidth + medianCellW)), CInt((offsetY + y * medianCellH + medianCellH)))
                    Dim down As Point = New Point(CInt((offsetX + x * cellWidth + cellWidth)), CInt((offsetY + y * medianCellH + cellHeight)))
                    renderCell(left, top, right, down, count, Color.DarkGray)
                    count += 1
                Next
            End If
        Next
    End Sub

    Private Sub renderCell(left As Point, top As Point, right As Point, down As Point, count As Integer, color As Color)
        Dim cell As New DiamondPictureBox()
        cell.BackColor = color
        cell.Size = New Size(36, 18)
        cell.Location = left
        cell.Name = count
        AddHandler cell.Click, Sub(e_sender, eventargs)
                                   MsgBox(e_sender.Name)
                               End Sub
        AddHandler cell.MouseMove, Sub(e_sender, eventargs)
                                       cell.BackColor = Color.Teal
                                   End Sub
        AddHandler cell.MouseLeave, Sub(e_sender, eventargs)
                                        cell.BackColor = Color.DarkGray
                                    End Sub
        IsometricGrid.Controls.Add(cell)
    End Sub

    Private Sub disableDrag()
        Drag = False
    End Sub

    Private Sub enableDrag()
        Drag = True
        Xaxys = Windows.Forms.Cursor.Position.X - Me.Left
        Yaxys = Windows.Forms.Cursor.Position.Y - Me.Top
    End Sub

    Private Sub moveControl()
        If Drag Then
            Me.Top = Windows.Forms.Cursor.Position.Y - Yaxys
            Me.Left = Windows.Forms.Cursor.Position.X - Xaxys
        End If
    End Sub

End Class
