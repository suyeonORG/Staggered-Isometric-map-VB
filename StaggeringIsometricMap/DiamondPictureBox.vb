Imports System.Drawing.Drawing2D
Public Class DiamondPictureBox
    Inherits PictureBox

    Protected Overrides Sub OnPaint(ByVal pe As PaintEventArgs)
        Using p = New GraphicsPath()
            p.AddPolygon(New Point() {New Point(Width / 2, 0), New Point(Width, Height / 2), New Point(Width / 2, Height), New Point(0, Height / 2)})
            Me.Region = New Region(p)
            MyBase.OnPaint(pe)
        End Using
    End Sub
End Class
