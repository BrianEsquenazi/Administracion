Public Class FormOrganizer

    Private form As Form
    Private maxHeight As Integer
    Private width As Integer
    Private controls As List(Of CustomControl)
    Private buttons As New List(Of CustomButton)
    Private buttonsTop As New List(Of CustomButton)
    Private buttonsBottom As New List(Of CustomButton)
    Private btnAdd As CustomButton
    Private btnAddClick As EventHandler
    Private btnDelete As CustomButton
    Private btnDeleteClick As EventHandler
    Private btnClean As CustomButton
    Private btnCleanClick As EventHandler
    Private btnQuery As CustomButton
    Private btnQueryClick As EventHandler
    Private btnList As CustomButton
    Private btnListClick As EventHandler
    Private btnClose As CustomButton
    Private btnCloseClick As EventHandler

    Private topMargin As Integer = 30
    Private leftMargin As Integer = 30
    Private rightMargin As Integer = 30
    Private bottomMargin As Integer = 30
    Private separation As Integer = 6
    Private simpleButtonHeight As Integer = 35
    Private separationBetweenControlsAndButtons As Integer = 45
    Private charPixelSize As Double = 7.5

    Public Delegate Sub ControlEvent(ByVal sender As Object, ByVal e As EventArgs)

    Public Sub New(ByVal someForm As Form, ByVal formWidth As Integer, ByVal formHeight As Integer)
        form = someForm
        maxHeight = formHeight
        width = formWidth
    End Sub

    Public Sub addControls(ByVal formControls As List(Of CustomControl))
        controls = formControls
        controls.OrderBy(Function(control) control.LabelAssociationKey)
    End Sub

    Private Function controlsHeight()
        Return controls.Count * 20 + (controls.Count - 1) * topMargin 'TODO control.getHeight() + separation (en realidad separation es 10 y no 6 acá)
    End Function

    Private Function buttonsHeight()
        Return simpleButtonHeight * 2 + separation
    End Function

    Public Sub organize()
        form.Height = Math.Min(maxHeight, topMargin + bottomMargin + controlsHeight() + separationBetweenControlsAndButtons + buttonsHeight())
        form.Width = width

        Dim btnsTop As Integer = organizeControls() + separationBetweenControlsAndButtons
        organizeButtons(btnsTop)
    End Sub

    Private Function organizeControls()
        Dim top As Integer = topMargin
        Dim left As Integer = leftMargin + separation + maxLabelWidth()

        'Setteo el top y el left de los controls y labels
        For Each control As CustomControl In controls
            control.setTop(top - 3)
            control.setLeft(left)

            Dim label As CustomLabel = labelFor(control.LabelAssociationKey)
            label.Top = top
            label.Left = leftMargin
            top += topMargin 'TODO control.getHeight() + separation (en realidad separation es 10 y no 6 acá)
        Next

        'Settep el width variable de los text box
        For Each textBox As CustomTextBox In controls.OfType(Of CustomTextBox)()
            Dim textWidth As Integer = Math.Min(Math.Round(textBox.MaxLength * charPixelSize), width - leftMargin - rightMargin - labelFor(textBox.LabelAssociationKey).Width - separation)
            textBox.setWidth(textWidth)
        Next

        Return top - topMargin
    End Function

    Private Function maxLabelWidth()
        Return controls.ConvertAll(Function(control) labelFor(control.LabelAssociationKey).Width).Max()
    End Function

    Private Sub organizeButtons(ByVal top As Integer)
        createButtons()

        Dim buttonWidth As Integer = (width - leftMargin - rightMargin - separation * 3) \ 3
        buttons.ForEach(Sub(button) button.Width = buttonWidth)
        buttons.ForEach(Sub(button) button.Height = simpleButtonHeight)

        Dim left As Integer = leftMargin
        For Each button As CustomButton In buttonsTop
            setButtonPosition(button, top, left)
            left += buttonWidth + separation
        Next

        left = 30
        For Each button As CustomButton In buttonsBottom
            setButtonPosition(button, top + button.Height + separation, left)
            left += buttonWidth + separation
        Next
    End Sub

    Private Sub setButtonPosition(ByVal button As CustomButton, ByVal top As Integer, ByVal left As Integer)
        button.Top = top
        button.Left = left
    End Sub

    Private Sub createButtons()
        buttonsTop.Add(addButton)
        buttonsTop.Add(deleteButton)
        buttonsTop.Add(cleanButton)

        buttonsBottom.Add(queryButton)
        buttonsBottom.Add(listButton)
        buttonsBottom.Add(closeButton)

        buttons.AddRange(buttonsTop)
        buttons.AddRange(buttonsBottom)
    End Sub

    Private Function addButton()
        Dim btn As CustomButton
        If IsNothing(btnAdd) Then
            btn = New CustomButton()
            btn.Parent = form
            btn.Name = "btnAdd"
            btn.Text = "Agregar"
        Else
            btn = btnAdd
        End If

        If Not IsNothing(btnAddClick) Then
            AddHandler btn.Click, btnAddClick
        End If

        Return btn
    End Function

    Private Function deleteButton()
        Dim btn As CustomButton
        If IsNothing(btnDelete) Then
            btn = New CustomButton()
            btn.Parent = form
            btn.Name = "btnDelete"
            btn.Text = "Eliminar"
        Else
            btn = btnDelete
        End If

        If Not IsNothing(btnDeleteClick) Then
            AddHandler btn.Click, btnDeleteClick
        End If

        Return btn
    End Function

    Private Function cleanButton()
        Dim btn As CustomButton
        If IsNothing(btnClean) Then
            btn = New CustomButton()
            btn.Parent = form
            btn.Name = "btnClean"
            btn.Text = "Limpiar"
        Else
            btn = btnClean
        End If

        If Not IsNothing(btnCleanClick) Then
            AddHandler btn.Click, btnCleanClick
        End If

        Return btn
    End Function

    Private Function queryButton()
        Dim btn As CustomButton
        If IsNothing(btnQuery) Then
            btn = New CustomButton()
            btn.Parent = form
            btn.Name = "btnQuery"
            btn.Text = "Consulta"
        Else
            btn = btnQuery
        End If

        If Not IsNothing(btnQueryClick) Then
            AddHandler btn.Click, btnQueryClick
        End If

        Return btn
    End Function

    Private Function listButton()
        Dim btn As CustomButton
        If IsNothing(btnList) Then
            btn = New CustomButton()
            btn.Parent = form
            btn.Name = "btnList"
            btn.Text = "Listado"
        Else
            btn = btnList
        End If

        If Not IsNothing(btnListClick) Then
            AddHandler btn.Click, btnListClick
        End If

        Return btn
    End Function

    Private Function closeButton()
        Dim btn As CustomButton
        If IsNothing(btnClose) Then
            btn = New CustomButton()
            btn.Parent = form
            btn.Name = "btnClose"
            btn.Text = "Cerrar"
        Else
            btn = btnClose
        End If

        If Not IsNothing(btnCloseClick) Then
            AddHandler btn.Click, btnCloseClick
        End If

        Return btn
    End Function

    Private Function labelFor(ByVal index As Integer) As CustomLabel
        Return form.Controls.OfType(Of CustomLabel).ToList.Find(Function(label) label.ControlAssociationKey = index)
    End Function

    Public Sub addButton(ByVal button As CustomButton)
        btnAdd = button
    End Sub
    Public Sub deleteButton(ByVal button As CustomButton)
        btnDelete = button
    End Sub
    Public Sub cleanButton(ByVal button As CustomButton)
        btnClean = button
    End Sub
    Public Sub queryButton(ByVal button As CustomButton)
        btnQuery = button
    End Sub
    Public Sub listButton(ByVal button As CustomButton)
        btnList = button
    End Sub
    Public Sub closeButton(ByVal button As CustomButton)
        btnClose = button
    End Sub
    Public Sub setAddButtonClick(ByRef btnClick As EventHandler)
        btnAddClick = btnClick
    End Sub
    Public Sub setDeleteButtonClick(ByRef btnClick As EventHandler)
        btnDeleteClick = btnClick
    End Sub
    Public Sub setCleanButtonClick(ByRef btnClick As EventHandler)
        btnCleanClick = btnClick
    End Sub
    Public Sub setQueryButtonClick(ByRef btnClick As EventHandler)
        btnQueryClick = btnClick
    End Sub
    Public Sub setListButtonClick(ByRef btnClick As EventHandler)
        btnListClick = btnClick
    End Sub
    Public Sub setCloseButtonClick(ByRef btnClick As EventHandler)
        btnCloseClick = btnClick
    End Sub
End Class
