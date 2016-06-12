Public Class CommonEventsHandler
    Private Shared controls As List(Of CustomControl)
    Private Shared isCRUDForm As Boolean = True

    Public Shared Sub setIndexTab(ByVal form As Form)
        controls = New List(Of CustomControl)
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

        For Each validableControl As CustomTextBox In form.Controls.OfType(Of CustomTextBox)() 'Ver si agregar los combo
            addValidableControlFormatTo(validableControl)
        Next

        form.Show()
        If Not IsNothing(firstControl) Then
            firstControl.Focus()
        End If
    End Sub

    Public Shared Sub setIndexTabNotCRUDForm(ByVal form As Form)
        isCRUDForm = False
        setIndexTab(form)
    End Sub

    Public Shared Sub addValidableControlFormatTo(ByVal validableControl As ValidableControl)
        Dim control As Control = DirectCast(validableControl, Control)
        Select Case validableControl.Validator
            Case ValidatorType.Numeric
                AddHandler control.KeyPress, AddressOf numericKeyPressed
            Case ValidatorType.Positive, ValidatorType.PositiveWithMax
                AddHandler control.KeyPress, AddressOf numericKeyOrDecimalSeparatorPressed
            Case ValidatorType.DateFormat
                AddHandler control.KeyDown, AddressOf deleteOrBackSpaceDownForDateFormat
                AddHandler control.KeyPress, AddressOf dateKeyPressed
                control.Text = "  /  /    "
        End Select
    End Sub

    Private Shared Sub numericKeyPressed(ByVal sender As Object, ByVal e As KeyPressEventArgs)
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Shared Sub numericKeyOrDecimalSeparatorPressed(ByVal sender As Object, ByVal e As KeyPressEventArgs)
        Dim customText = DirectCast(sender, CustomTextBox)
        If e.KeyChar = "." Or e.KeyChar = "," Then
            If Not customText.Text.Contains(".") Then
                customText.Text = customText.Text.Insert(customText.Text.Count, ".")
                customText.Select(customText.TextLength, 0)
            End If
            e.Handled = True
        End If
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Shared Sub deleteOrBackSpaceDownForDateFormat(ByVal sender As Object, ByVal e As KeyEventArgs)
        Dim customControl As CustomTextBox = DirectCast(sender, CustomTextBox)
        If e.KeyValue = Keys.Delete Then
            deleteCharOf(customControl, customControl.SelectionStart, 1)
            e.SuppressKeyPress = True
        End If
        If e.KeyValue = Keys.Back Then
            deleteCharOf(customControl, customControl.SelectionStart - 1, 0)
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Shared Sub deleteCharOf(ByVal textBox As CustomTextBox, ByVal selectionStart As Integer, ByVal selectionNextPosition As Integer)
        Select Case selectionStart
            Case 0, 1, 3, 4, 6 To 9
                textBox.Text = textBox.Text.Remove(selectionStart, 1).Insert(selectionStart, " ")
                Dim nextSelection As Integer = Math.Max(selectionStart + selectionNextPosition, selectionNextPosition)
                If nextSelection = 10 Then
                    textBox.Select(0, 0)
                Else
                    textBox.Select(nextSelection, 0)
                End If
            Case 2, 5 'No se borra nada, porque es un slash
                textBox.Select(Math.Max(selectionStart + selectionNextPosition, 1), 0)
        End Select
    End Sub

    Private Shared Sub dateKeyPressed(ByVal sender As Object, ByVal e As KeyPressEventArgs)
        Dim customControl As CustomTextBox = DirectCast(sender, CustomTextBox)
        Dim firstSpaceIndex As Integer = customControl.Text.IndexOf(" ")
        Select Case firstSpaceIndex
            Case 0, 1, 3, 4, 6 To 9 'Es una entrada de fecha
                If IsNumeric(e.KeyChar) Then
                    customControl.Text = customControl.Text.Remove(firstSpaceIndex, 1).Insert(firstSpaceIndex, e.KeyChar)
                    customControl.Select(Math.Max(customControl.Text.IndexOf(" "), 0), 0)
                    e.Handled = True
                End If
        End Select
    End Sub

    Private Shared Sub enterPressed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            Dim nextControl As Control = controls.Find(Function(control) control.EnterIndex = sender.EnterIndex + 1)
            If IsNothing(nextControl) OrElse Not nextControl.Focus() Then
                If isCRUDForm Then
                    secondControl().Focus()
                Else
                    firstControl.Focus()
                End If
            End If
        End If
    End Sub

    Private Shared Function secondControl() As Control
        Dim sndControl As Control
        sndControl = controls.Find(Function(control) control.EnterIndex = 2)
        Return sndControl
    End Function

    Private Shared Function firstControl() As Control
        Dim fstControl As Control
        fstControl = controls.Find(Function(control) control.EnterIndex = 1)
        Return fstControl
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
