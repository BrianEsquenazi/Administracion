Public Delegate Function QueryFunction(ByVal text As String)
Public Delegate Sub ShowMethod(ByVal selectedValue)

Public Class FormOrganizer

    Private form As Form
    Private maxHeight As Integer
    Private width As Integer
    Private queryFunctions As New List(Of Tuple(Of QueryFunction, String, ShowMethod))
    Private queryFunction As QueryFunction
    Private queryText As CustomTextBox
    Private queryList As CustomListBox
    Private selectionList As CustomListBox
    Private controls As List(Of CustomControl)
    Private annexedControls As New List(Of CustomControl)
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
    Private showMethodFunction As ShowMethod
    Private usingQueryText As Boolean

    Private topMargin As Integer = 30
    Private leftMargin As Integer = 30
    Private rightMargin As Integer = 30
    Private bottomMargin As Integer = 30
    Private separation As Integer = 6
    Private simpleButtonHeight As Integer = 35
    Private listQueryHeight As Integer = 240
    Private separationBetweenControlsAndButtons As Integer = 45
    Private controlSeparation As Integer = 10
    Private charPixelSize As Double = 7.5

    Public Sub New(ByVal someForm As Form, ByVal formWidth As Integer, ByVal formHeight As Integer)
        form = someForm
        maxHeight = formHeight
        width = formWidth
    End Sub

    Public Sub addControls(ByVal formControls As List(Of CustomControl))
        controls = formControls
        controls.OrderBy(Function(control) control.LabelAssociationKey)
    End Sub

    Public Sub addAnnexedControls(ByVal someControls As List(Of CustomControl))
        annexedControls = someControls
        annexedControls.OrderBy(Function(control) control.LabelAssociationKey)
    End Sub

    Private Function controlsHeight()
        Return controls.Sum(Function(control As CustomControl) DirectCast(control, Control).Height) + (controls.Count - 1) * controlSeparation
    End Function

    Private Function buttonsHeight()
        Return simpleButtonHeight * 2 + separation 'Las dos filas de botones + separation entre botones
    End Function

    Private Function formNormalHeight()
        Return topMargin + bottomMargin + controlsHeight() + separationBetweenControlsAndButtons + buttonsHeight()
    End Function

    Private Function formWithQueryControlsHeight()
        Dim height As Integer = formNormalHeight() + separation + queryList.Height
        If usingQueryText Then
            height += separation + queryTextHeight()
        End If
        Return height
    End Function

    Private Function formWithSelectionListHeight()
        Return formNormalHeight() + separation + selectionList.Height
    End Function

    Public Sub organize()
        CommonEventsHandler.setIndexTab(form)
        form.Height = Math.Min(maxHeight, formNormalHeight)
        form.Width = width

        Dim btnsTop As Integer = organizeControls() + separationBetweenControlsAndButtons
        organizeButtons(btnsTop)
        organizeQueryControllers(btnsTop + simpleButtonHeight * 2 + separation)
    End Sub

    Private Function organizeControls()
        Dim top As Integer = topMargin
        Dim left As Integer = leftMargin + separation + maxLabelWidth()

        'Setteo el top y el left de los controls y labels
        For Each control As CustomControl In controls
            Dim castControl As Control = DirectCast(control, Control)
            castControl.Top = top - 3
            castControl.Left = left
            castControl.Width = variableWidthFor(control)

            Dim label As CustomLabel = labelFor(control.LabelAssociationKey)
            label.Top = top
            label.Left = leftMargin

            Dim annexedCustomControl As CustomControl = annexedControlFor(control.LabelAssociationKey)
            Dim annexedControl As Control = DirectCast(annexedCustomControl, Control)
            If Not IsNothing(annexedControl) Then
                annexedControl.Top = top - 3
                annexedControl.Left = left + castControl.Width + separation

                Dim annexedWidth As Integer
                If variableWidthFor(annexedCustomControl) - castControl.Width = 0 Then
                    castControl.Width = variableWidthFor(control) \ 2
                    annexedWidth = castControl.Width - 6
                Else
                    annexedWidth = variableWidthFor(annexedCustomControl) - castControl.Width - separation
                End If
                annexedControl.Width = annexedWidth
            End If
            top += castControl.Height + controlSeparation
        Next

        Return top - topMargin
    End Function

    Private Function variableWidthFor(ByVal textBox As CustomTextBox)
        Return Math.Min(Math.Round(textBox.MaxLength * charPixelSize), variableWidthOfControl())
    End Function

    Private Function variableWidthFor(ByVal control As CustomControl)
        Try
            Return variableWidthFor(DirectCast(control, CustomTextBox))
        Catch e As InvalidCastException
            Return variableWidthOfControl()
        End Try
    End Function

    Private Function variableWidthOfControl()
        Return width - leftMargin - rightMargin - maxLabelWidth() - separation
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

    Private Sub organizeQueryControllers(ByVal top As Integer)
        If usingQueryText Then
            queryText = New CustomTextBox
            queryText.Parent = form
            queryText.Name = "txtQuery"
            queryText.Width = form.Width - leftMargin - rightMargin
            queryText.Top = top + separation
            top = queryText.Top
            queryText.Left = leftMargin
            queryText.Visible = False
            AddHandler queryText.KeyDown, AddressOf queryTextEnterPressed
        End If

        queryList = New CustomListBox
        queryList.Parent = form
        queryList.Name = "lstQuery"
        queryList.Width = form.Width - leftMargin - rightMargin
        queryList.Height = listQueryHeight
        queryList.Top = top + queryTextHeight() + separation
        queryList.Left = leftMargin
        queryList.Visible = False
        AddHandler queryList.DoubleClick, AddressOf listDoubleClickEventWithHide

        If queryFunctions.Count > 1 Then
            selectionList = New CustomListBox
            selectionList.Parent = form
            selectionList.Name = "lstQuery"
            selectionList.Width = form.Width - leftMargin - rightMargin
            selectionList.Height = Math.Round(listQueryHeight / 2)
            selectionList.Top = top + separation
            selectionList.Left = leftMargin
            selectionList.Visible = False
            AddHandler selectionList.DoubleClick, AddressOf selectionDoubleClick
        End If
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
            AddHandler btn.Click, AddressOf addClickWithClean
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
            AddHandler btn.Click, AddressOf deleteClickWithConfirmation
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

    Private Function annexedControlFor(ByVal index As Integer) As CustomControl
        Return annexedControls.Find(Function(control) control.LabelAssociationKey = index)
    End Function

    Private Function validateForDelete()
        Return validateControls(False)
    End Function

    Private Function validateForAdd()
        Return validateControls(True)
    End Function

    Private Function validateControls(ByVal isAdd As Boolean)
        Dim firstControl As CustomTextBox = controls.Find(Function(control) control.EnterIndex = 1) 'TODO HACERLO GENÉRICO PARA TODOS LOS CONTROLERS
        Dim validator As New Validator
        validator.validate(firstControl.Text, firstControl.Validator, labelFor(firstControl.LabelAssociationKey).Text)
        If isAdd Then
            Dim controlsToValidate As List(Of CustomTextBox) = controls.OfType(Of CustomTextBox).ToList 'TODO HACERLO GENÉRICO PARA TODOS LOS CONTROLERS
            controlsToValidate.Remove(firstControl)
            controlsToValidate.ForEach(Sub(control) validator.validate(control.Text, control.Validator, labelFor(control.LabelAssociationKey).Text))
        End If
        Return validator.flush()
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

    Public Sub setDefaultCleanButtonClick()
        btnCleanClick = AddressOf defaultCleanClick
    End Sub
    Public Sub addQueryFunction(ByVal listFunction As QueryFunction, ByVal name As String, ByVal showFunction As ShowMethod)
        queryFunctions.Add(Tuple.Create(listFunction, name, showFunction))
        btnQueryClick = AddressOf defaultQueryClick
    End Sub
    Public Sub dontUseQueryText()
        usingQueryText = False
    End Sub
    Public Sub setDefaultCloseButtonClick()
        btnCloseClick = AddressOf defaultCloseClick
    End Sub

    Private Sub listDoubleClickEventWithHide(ByVal sender As Object, ByVal e As EventArgs)
        showMethodFunction.Invoke(queryList.SelectedValue)
        queryFunction = Nothing
        hideQueryControls()
    End Sub
    Private Sub selectionDoubleClick(ByVal sender As Object, ByVal e As EventArgs)
        queryFunction = selectionList.SelectedValue.Item1
        showMethodFunction = selectionList.SelectedValue.Item3
        hideSelectionList()
        showQueryList()
    End Sub

    Private Sub addClickWithClean(ByVal sender As Object, ByVal e As EventArgs)
        If validateForAdd() Then
            btnAddClick.Invoke(sender, e)
            Cleanner.clean(form)
        End If
    End Sub
    Private Sub deleteClickWithConfirmation(ByVal sender As Object, ByVal e As EventArgs)
        If validateForDelete() Then
            If MsgBox("¿Desea eliminar el registro?", MsgBoxStyle.YesNo, "Eliminar") = vbYes Then
                btnDeleteClick.Invoke(sender, e)
                Cleanner.clean(form)
            End If
        End If
    End Sub

    Private Sub defaultCleanClick(ByVal sender As Object, ByVal e As EventArgs)
        Cleanner.clean(form)
    End Sub
    Private Sub defaultQueryClick(ByVal sender As Object, ByVal e As EventArgs)
        Select Case queryFunctions.Count
            Case 0
            Case 1
                queryFunction = queryFunctions.First().Item1
                showMethodFunction = queryFunctions.First().Item3
                showQueryList()
            Case Else
                If queryList.Visible Then
                    hideQueryControls()
                End If
                showSelectionList()
        End Select
    End Sub
    Private Sub showQueryList()
        queryTextEnterPressed(Nothing, Nothing)
        If usingQueryText Then
            queryText.Visible = True
            queryText.Focus()
        End If
        queryList.Visible = True
        form.Height = formWithQueryControlsHeight()
    End Sub
    Private Sub showSelectionList()
        selectionList.DataSource = queryFunctions
        selectionList.DisplayMember = "Item2"
        selectionList.Visible = True
        form.Height = formWithSelectionListHeight()
    End Sub

    Private Sub defaultCloseClick(ByVal sender As Object, ByVal e As EventArgs)
        form.Close()
    End Sub

    Private Sub queryTextEnterPressed(ByVal sender As Object, ByVal e As EventArgs)
        Dim searchText As String = ""
        If usingQueryText Then
            searchText = queryText.Text
        End If
        queryList.DataSource = queryFunction.Invoke(searchText)
    End Sub

    Private Sub hideQueryControls()
        If usingQueryText Then
            queryText.Visible = False
        End If
        queryList.Visible = False
        hideSelectionList()
        form.Height = formNormalHeight()
    End Sub
    Private Sub hideSelectionList()
        If Not IsNothing(selectionList) Then
            selectionList.Visible = False
            form.Height = formNormalHeight()
        End If
    End Sub

    Private Function queryTextHeight() As Integer
        Dim heigth As Integer = 0
        If usingQueryText Then
            heigth = queryText.Height
        End If
        Return heigth
    End Function

End Class
