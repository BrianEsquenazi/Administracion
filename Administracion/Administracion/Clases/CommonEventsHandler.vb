Public Class CommonEventsHandler
    Private Shared controls As New List(Of CustomControl)

    Public Shared Sub setIndexTab(ByVal form As Form)
        For Each txtBox As CustomTextBox In form.Controls.OfType(Of CustomTextBox)()
            If txtBox.EnterIndex > 0 Then
                AddHandler txtBox.KeyDown, AddressOf enterPressed
                AddHandler txtBox.KeyPress, AddressOf enterOrEscapePressedWithoutSound
                controls.Add(txtBox)
            End If
        Next
        For Each cmbBox As CustomComboBox In form.Controls.OfType(Of CustomComboBox)()
            If cmbBox.EnterIndex > 0 Then
                AddHandler cmbBox.KeyDown, AddressOf enterPressed
                controls.Add(cmbBox)
            End If
        Next
        For Each lstBox As CustomListBox In form.Controls.OfType(Of CustomListBox)()
            If lstBox.EnterIndex > 0 Then
                AddHandler lstBox.KeyDown, AddressOf enterPressed
                controls.Add(lstBox)
            End If
        Next
        For Each btn As CustomButton In form.Controls.OfType(Of CustomButton)()
            If btn.EnterIndex > 0 Then
                AddHandler btn.KeyDown, AddressOf enterPressed
                controls.Add(btn)
            End If
        Next

        Dim firstControl As Control
        firstControl = controls.Find(Function(control) control.EnterIndex = 1)
        form.Show()
        If Not IsNothing(firstControl) Then
            firstControl.Focus()
        End If
    End Sub

    Private Shared Sub enterPressed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            Dim nextControl As Control = controls.Find(Function(control) control.EnterIndex = sender.EnterIndex + 1)
            If IsNothing(nextControl) OrElse Not nextControl.Focus() Then
                secondControl().Focus()
            End If
        End If
    End Sub

    Private Shared Function secondControl() As Control
        Dim sndControl As Control
        sndControl = controls.Find(Function(control) control.EnterIndex = 2)
        Return sndControl
    End Function

    Private Shared Sub enterOrEscapePressedWithoutSound(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Asc(e.KeyChar) = Keys.Enter Then
            e.Handled = True
        End If
        If Asc(e.KeyChar) = Keys.Escape Then
            DirectCast(sender, CustomTextBox).Text = ""
            e.Handled = True
        End If
    End Sub
End Class
