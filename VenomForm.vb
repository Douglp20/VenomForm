Imports System.Windows.Forms


Public Class VenomForm
    Private VenomRegistry As New Douglas.Venom.Registry.VenomRegistry
    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)
    Public Sub New()
    End Sub
    Public Sub Unload(ByRef frm As System.Windows.Forms.Form)
        On Error GoTo Err

        VenomRegistry.SaveSetting("FORMSETTING", frm.Name + "Top", frm.Top)
        VenomRegistry.SaveSetting("FORMSETTING", frm.Name + "Left", frm.Left)
        VenomRegistry.SaveSetting("FORMSETTING", frm.Name + "Width", frm.Width)
        VenomRegistry.SaveSetting("FORMSETTING", frm.Name + "Height", frm.Height)

        SetControlProperties(frm, frm.Name.ToString, "unload")


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Sub
    Public Sub Load(ByRef frm As System.Windows.Forms.Form)
        On Error GoTo Err


        Dim top As String = VenomRegistry.GetSetting("FORMSETTING", frm.Name + "Top", "0")
        Dim Left As String = VenomRegistry.GetSetting("FORMSETTING", frm.Name + "Left", "0")
        Dim width As String = VenomRegistry.GetSetting("FORMSETTING", frm.Name + "Width", "0")
        Dim height As String = VenomRegistry.GetSetting("FORMSETTING", frm.Name + "Height", "0")

        If width = frm.Width Then
            top = 3
            Left = 3
            'width = frm.Width + 200
            'height = frm.Height

            'frm.FormBorderStyle = FormBorderStyle.Sizable
            frm.Top = top.ToString()
            frm.Left = Left.ToString()
            frm.Height = height.ToString()
            frm.Width = width.ToString()
        End If
        If Not width = frm.Width Then
            top = 3
            Left = 3
            'width = frm.Width + 200
            'height = frm.Height
            frm.FormBorderStyle = FormBorderStyle.Sizable
        End If
        If frm.FormBorderStyle = FormBorderStyle.FixedSingle Then
            frm.Top = top.ToString()
            frm.Left = Left.ToString()
        End If

        If frm.FormBorderStyle = FormBorderStyle.Sizable Then
            frm.Top = top.ToString()
            frm.Left = Left.ToString()
            frm.Height = height.ToString()
            frm.Width = width.ToString()
        End If
        SetControlProperties(frm, frm.Name.ToString, "load")


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Sub


    Private Sub SetControlProperties(control As System.Windows.Forms.Control, formName As String, action As String)
        On Error GoTo Err
        If control.GetType.ToString = "System.Windows.Forms.DataGridView" Then
            Dim dgv As New System.Windows.Forms.DataGridView
            dgv = CType(control, System.Windows.Forms.DataGridView)
            Dim strKey As String = formName + dgv.Name.ToString
            Dim col As Integer = dgv.Columns.Count
            Dim widthsize As String
            Dim iwidth As Integer

            If action = "load" Then
                For c As Integer = 0 To dgv.Columns.Count - 1
                    widthsize = VenomRegistry.GetSetting("DATAGRIDVIEW", strKey + "col" + c.ToString, 60)
                    If c = 0 Then
                        If dgv.Columns(c).HeaderText = "ID" Then
                            dgv.Columns(c).Visible = False
                        Else
                            dgv.Columns(c).Width = widthsize.ToString
                        End If
                    Else
                        dgv.Columns(c).Width = widthsize.ToString
                    End If
                Next
            Else
                For c As Integer = 0 To col - 1
                    iwidth = dgv.Columns(c).Width
                    VenomRegistry.SaveSetting("DATAGRIDVIEW", strKey + "col" + c.ToString, iwidth.ToString)

                Next
            End If
        End If

        If control.GetType.ToString = "System.Windows.Forms.ListView" Then
            Dim lvw As New System.Windows.Forms.ListView
            lvw = CType(control, System.Windows.Forms.ListView)
            Dim strKey As String = formName + lvw.Name.ToString
            Dim col As Integer = lvw.Columns.Count
            Dim widthsize As String
            Dim iwidth As Integer

            If action = "load" Then
                For i As Integer = 0 To col - 1

                    If i = 0 And lvw.CheckBoxes = True Then
                        If lvw.Columns.Count > 1 Then lvw.Columns(i).Width = 18
                    Else
                        If i = 0 And lvw.CheckBoxes = False Then
                            lvw.Columns(i).Width = 0
                        Else
                            widthsize = VenomRegistry.GetSetting("LISTVIEW", strKey + "col" + i.ToString, 60)

                            If lvw.Columns(i).Tag = "hide" Then
                                lvw.Columns(i).Width = 0
                            Else
                                lvw.Columns(i).Width = widthsize.ToString
                            End If
                        End If
                    End If

                Next
            Else
                For i As Integer = 0 To col - 1
                    iwidth = lvw.Columns(i).Width
                    VenomRegistry.SaveSetting("LISTVIEW", strKey + "col" + i.ToString, iwidth.ToString)

                Next
            End If

        End If
        'Go to the level of each control
        For Each subControl In control.Controls
            SetControlProperties(subControl, formName, action)
        Next
        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module (within Control : " + control.Name + ") " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub

    Public Sub DataLoad(ByRef frm As System.Windows.Forms.Form, ByRef ds As DataSet)
        On Error GoTo Err



        SetControlDataLoad(frm, frm.Name.ToString, ds)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
    Private Sub SetControlDataLoad(control As System.Windows.Forms.Control, formName As String, ByRef ds As DataSet)
        On Error GoTo Err

        Dim strFieldName As String
        Dim strDataField As String

        Select Case control.GetType.ToString
            Case "System.Windows.Forms.TextBox"

                Dim txt As New System.Windows.Forms.TextBox
                txt = CType(control, System.Windows.Forms.TextBox)
                Dim tagField As String
                Dim dataField As String
                Dim dataValue As String
                Dim dataType As String
                Dim txtNameTag As String
                Dim txtName As String
                ' Dim c As Integer = 0

                If ds.Tables(0).Rows.Count > 0 Then
                    For Each row As DataRow In ds.Tables(0).Rows
                        For c As Integer = 0 To row.ItemArray.Count - 1
                            ' dataType = ds.Tables(0).Rows(0).Item(c).GetType().ToString
                            dataType = row.Table.Columns(c).DataType().ToString
                            dataField = row.Table.Columns(c).ColumnName
                            strDataField = dataField
                            txtNameTag = txt.Tag
                            txtName = txt.Name
                            strFieldName = txtName.ToString

                            If Len(txtNameTag) > 0 Then
                                If dataField.ToLower() = txtNameTag.ToLower() Then
                                    Select Case dataType.ToString
                                        Case "System.String"
                                            txt.Text = row(txtNameTag).ToString
                                        Case Else
                                            txt.Text = CType(row(txtNameTag).ToString, String)

                                    End Select
                                End If
                            End If
                        Next
                    Next
                End If
            Case "System.Windows.Forms.MaskedTextBox"


                Dim msk As New System.Windows.Forms.MaskedTextBox
                msk = CType(control, System.Windows.Forms.MaskedTextBox)
                Dim tagField As String
                Dim dataField As String
                Dim dataValue As String
                Dim dataType As String
                Dim mskNameTag As String
                Dim mskName As String
                ' Dim c As Integer = 0

                If ds.Tables(0).Rows.Count > 0 Then
                    For Each row As DataRow In ds.Tables(0).Rows
                        For c As Integer = 0 To row.ItemArray.Count - 1
                            dataType = row.Table.Columns(c).DataType().ToString
                            dataField = row.Table.Columns(c).ColumnName
                            strDataField = dataField
                            mskNameTag = msk.Tag
                            mskName = msk.Name
                            strFieldName = mskName.ToString

                            If Len(mskNameTag) > 0 Then
                                If dataField.ToLower() = mskNameTag.ToLower() Then
                                    Select Case dataType.ToString
                                        Case "System.String"
                                            msk.Text = row(mskNameTag).ToString
                                        Case Else
                                            msk.Text = CType(row(mskNameTag).ToString, String)

                                    End Select
                                End If
                            End If
                        Next
                    Next
                End If

            Case "System.Windows.Forms.RichTextBox"

                Dim rtx As New System.Windows.Forms.RichTextBox
                rtx = CType(control, System.Windows.Forms.RichTextBox)
                Dim tagField As String
                Dim dataField As String
                Dim dataValue As String
                Dim dataType As String
                Dim rtxNameTag As String
                Dim rtxName As String
                ' Dim c As Integer = 0

                If ds.Tables(0).Rows.Count > 0 Then
                    For Each row As DataRow In ds.Tables(0).Rows
                        For c As Integer = 0 To row.ItemArray.Count - 1
                            ' dataType = ds.Tables(0).Rows(0).Item(c).GetType().ToString
                            dataType = row.Table.Columns(c).DataType().ToString
                            dataField = row.Table.Columns(c).ColumnName
                            strDataField = dataField
                            rtxNameTag = rtx.Tag
                            rtxName = rtx.Name
                            strFieldName = rtxName.ToString

                            If Len(rtxNameTag) > 0 Then
                                If dataField.ToLower() = rtxNameTag.ToLower() Then
                                    Select Case dataType.ToString
                                        Case "System.String"
                                            rtx.Text = row(rtxNameTag).ToString
                                        Case Else
                                            rtx.Text = CType(row(rtxNameTag).ToString, String)

                                    End Select
                                End If
                            End If
                        Next
                    Next
                End If
            Case "System.Windows.Forms.ComboBox"
                Dim cbo As New System.Windows.Forms.ComboBox
                cbo = CType(control, System.Windows.Forms.ComboBox)
                Dim tagField As String
                Dim dataField As String
                Dim dataValue As String
                Dim dataType As String
                Dim cboNameTag As String
                Dim cboName As String
                ' Dim c As Integer = 0

                If ds.Tables(0).Rows.Count > 0 Then
                    For Each row As DataRow In ds.Tables(0).Rows
                        For c As Integer = 0 To row.ItemArray.Count - 1
                            ' dataType = ds.Tables(0).Rows(0).Item(c).GetType().ToString
                            dataType = row.Table.Columns(c).DataType().ToString
                            dataField = row.Table.Columns(c).ColumnName
                            strDataField = dataField
                            cboNameTag = cbo.Tag
                            cboName = cbo.Name
                            strFieldName = cboName.ToString
                            Dim bolCheck As Boolean
                            If Len(cboNameTag) > 0 Then
                                If dataField.ToLower() = cboNameTag.ToLower() Then
                                    Select Case dataType.ToString
                                        Case "System.String"
                                            '  cbo.SelectedIndex = 0
                                            cbo.Text = row(cboNameTag).ToString
                                    End Select
                                End If
                            End If
                        Next
                    Next
                End If
            Case "System.Windows.Forms.CheckBox"
                Dim chk As New System.Windows.Forms.CheckBox
                chk = CType(control, System.Windows.Forms.CheckBox)
                Dim tagField As String
                Dim dataField As String
                Dim dataValue As String
                Dim dataType As String
                Dim chkNameTag As String
                Dim chkName As String
                ' Dim c As Integer = 0

                If ds.Tables(0).Rows.Count > 0 Then
                    For Each row As DataRow In ds.Tables(0).Rows
                        For c As Integer = 0 To row.ItemArray.Count - 1
                            ' dataType = ds.Tables(0).Rows(0).Item(c).GetType().ToString
                            dataType = row.Table.Columns(c).DataType().ToString
                            dataField = row.Table.Columns(c).ColumnName
                            strDataField = dataField
                            chkNameTag = chk.Tag
                            chkName = chk.Name
                            strFieldName = chkName.ToString
                            Dim bolCheck As Boolean
                            If Len(chkNameTag) > 0 Then
                                If dataField.ToLower() = chkNameTag.ToLower() Then
                                    Select Case dataType.ToString
                                        Case "System.Boolean", "System.Bit"
                                            chk.Checked = CBool(row(chkNameTag).ToString)
                                    End Select
                                End If
                            End If
                        Next
                    Next
                End If

            Case "System.Windows.Forms.DateTimePicker"
                Dim dte As New System.Windows.Forms.DateTimePicker
                dte = CType(control, System.Windows.Forms.DateTimePicker)
                Dim tagField As String
                Dim dataField As String
                Dim dataValue As String
                Dim dataType As String
                Dim dteNameTag As String
                Dim dteName As String
                ' Dim c As Integer = 0

                If ds.Tables(0).Rows.Count > 0 Then
                    For Each row As DataRow In ds.Tables(0).Rows
                        For c As Integer = 0 To row.ItemArray.Count - 1
                            ' dataType = ds.Tables(0).Rows(0).Item(c).GetType().ToString
                            dataType = row.Table.Columns(c).DataType().ToString
                            dataField = row.Table.Columns(c).ColumnName
                            strDataField = dataField
                            dteNameTag = dte.Tag
                            dteName = dte.Name
                            strFieldName = dteName.ToString
                            Dim bolCheck As Boolean

                            If Len(dteNameTag) > 0 Then
                                If dataField.ToLower = dteNameTag.ToLower() Then
                                    Select Case dataType.ToString
                                        Case "System.DateTime"
                                            bolCheck = IsDate(row(dteNameTag).ToString)
                                            If bolCheck Then dte.Value = row(dteNameTag).ToString
                                            dte.Checked = bolCheck
                                    End Select
                                End If
                            End If
                        Next
                    Next
                End If
            Case "System.Windows.Forms.PictureBox"

                Dim pic As New System.Windows.Forms.PictureBox
                pic = CType(control, System.Windows.Forms.PictureBox)
                pic.Image = Nothing
                Dim picField As String
                Dim dataField As String
                Dim dataValue As String
                Dim dataType As String
                Dim picName As String = pic.Name
                Dim picNameTag As String = pic.Tag



                If ds.Tables(0).Rows.Count > 0 Then
                    Dim converter As New System.Drawing.ImageConverter
                    Dim dataPicture As Byte()

                    For Each row As DataRow In ds.Tables(0).Rows
                        For c As Integer = 0 To row.ItemArray.Count - 1

                            dataType = row.Table.Columns(c).DataType().ToString
                            dataField = row.Table.Columns(c).ColumnName
                            strDataField = dataField
                            picNameTag = pic.Tag
                            picName = pic.Name
                            strFieldName = picName.ToString
                            If Len(picNameTag) > 0 Then
                                Select Case dataType.ToString()
                                    Case "System.Byte[]"
                                        If Not String.IsNullOrEmpty(row(dataField).ToString()) Then
                                            Dim DataImage As Byte() = DirectCast(row(dataField), Byte())
                                            If DataImage.Length > 0 Then
                                                Dim msPicture As System.IO.MemoryStream = New System.IO.MemoryStream(DataImage)
                                                pic.Image = System.Drawing.Image.FromStream(msPicture)
                                            End If
                                        End If
                                End Select
                            End If

                        Next
                    Next
                End If

            Case "System.Windows.Forms.GroupBox"

                Dim grp As New System.Windows.Forms.GroupBox
                grp = CType(control, System.Windows.Forms.GroupBox)
                Dim grpField As String
                Dim dataField As String
                Dim dataValue As String
                Dim dataType As String
                Dim grpName As String = grp.Name
                Dim grpNameTag As String = grp.Tag
                Dim rbtn As System.Windows.Forms.RadioButton

                If Len(grpNameTag) > 0 Then
                    If ds.Tables(0).Rows.Count > 0 Then
                        For Each row As DataRow In ds.Tables(0).Rows
                            For c As Integer = 0 To row.ItemArray.Count - 1
                                ' dataType = ds.Tables(0).Rows(0).Item(c).GetType().ToString
                                dataType = row.Table.Columns(c).DataType().ToString
                                dataField = row.Table.Columns(c).ColumnName
                                strDataField = dataField
                                strFieldName = grpName.ToString
                                Dim bolCheck As Boolean
                                Debug.WriteLine(dataField.ToLower.ToString)
                                If dataField.ToLower = grpNameTag.ToLower() Then
                                    Select Case dataType.ToString
                                        Case "System.Int", "System.Int32"
                                            Dim grpValue As Integer = CInt(row(grpNameTag).ToString)
                                            For Each ctlOpt In grp.Controls
                                                If TypeOf ctlOpt Is System.Windows.Forms.RadioButton Then
                                                    rbtn = DirectCast(ctlOpt, System.Windows.Forms.RadioButton)
                                                    If rbtn.Tag = grpValue Then
                                                        rbtn.Checked = True
                                                        Exit For
                                                    End If
                                                End If
                                            Next
                                    End Select
                                End If

                            Next
                        Next
                    End If

                End If
            Case "System.Windows.Forms.NumericUpDown"
                Dim npd As New System.Windows.Forms.NumericUpDown
                npd = CType(control, System.Windows.Forms.NumericUpDown)
                Dim tagField As String
                Dim dataField As String
                Dim dataValue As String
                Dim dataType As String
                Dim npdNameTag As String
                Dim npdName As String

                If ds.Tables(0).Rows.Count > 0 Then
                    For Each row As DataRow In ds.Tables(0).Rows
                        For c As Integer = 0 To row.ItemArray.Count - 1
                            ' dataType = ds.Tables(0).Rows(0).Item(c).GetType().ToString
                            dataType = row.Table.Columns(c).DataType().ToString
                            dataField = row.Table.Columns(c).ColumnName
                            strDataField = dataField
                            npdNameTag = npd.Tag
                            npdName = npd.Name
                            strFieldName = npdName.ToString
                            Dim bolCheck As Boolean

                            If Len(npdNameTag) > 0 Then
                                If dataField.ToLower = npdNameTag.ToLower() Then
                                    Select Case dataType.ToString
                                        Case "System.Int", "System.Int32"
                                            npd.Value = CInt(row(npdNameTag).ToString)

                                    End Select
                                End If
                            End If
                        Next
                    Next
                End If
        End Select





        'Go to the level of each control
        For Each subControl In control.Controls
            SetControlDataLoad(subControl, formName, ds)
        Next
        Exit Sub

Err:
        Dim rtn As String = "The error occur with " + control.GetType.ToString + " (the control name is " + strDataField + ") within the module" + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub

    Public Sub RefreshForm(ByRef frm As System.Windows.Forms.Form)
        On Error GoTo Err



        SetControlRefreshForm(frm, frm.Name.ToString)


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)


    End Sub
    Private Sub SetControlRefreshForm(control As System.Windows.Forms.Control, formName As String)
        On Error GoTo Err

        Dim strFieldName As String
        Select Case control.GetType.ToString
            Case "System.Windows.Forms.TextBox"
                Dim txt As New System.Windows.Forms.TextBox
                txt = CType(control, System.Windows.Forms.TextBox)
                If txt.Tag = "" Then
                Else
                    txt.Clear()
                End If
            Case "System.Windows.Forms.PictureBox"
                Dim pic As New System.Windows.Forms.PictureBox
                pic = CType(control, System.Windows.Forms.PictureBox)
                If pic.Tag = "" Then
                Else
                    pic.Image = Nothing
                End If
            Case "System.Windows.Forms.MaskedTextBox"
                Dim msk As New System.Windows.Forms.MaskedTextBox
                msk = CType(control, System.Windows.Forms.MaskedTextBox)
                If msk.Tag = "" Then
                Else
                    msk.Clear()
                End If
            Case "System.Windows.Forms.ComboBox"
                Dim cbo As New System.Windows.Forms.ComboBox
                cbo = CType(control, System.Windows.Forms.ComboBox)
                If cbo.Tag = "" Then
                Else
                    cbo.SelectedIndex = 0
                End If
            Case "System.Windows.Forms.CheckBox"
                Dim chk As New System.Windows.Forms.CheckBox
                chk = CType(control, System.Windows.Forms.CheckBox)
                If chk.Tag = "" Then
                Else
                    chk.Checked = False
                End If
            Case "System.Windows.Forms.NumericUpDown"
                Dim nup As New System.Windows.Forms.NumericUpDown
                nup = CType(control, System.Windows.Forms.NumericUpDown)
                If nup.Tag = "" Then
                Else
                    nup.Value = 0
                End If

            Case "System.Windows.Forms.DateTimePicker"
                Dim dte As New System.Windows.Forms.DateTimePicker
                dte = CType(control, System.Windows.Forms.DateTimePicker)
                If dte.Tag = "" Then
                Else
                    dte.Checked = False
                End If
            Case Else

        End Select


        'Go to the level of each control
        For Each subControl In control.Controls
            SetControlRefreshForm(subControl, formName)
        Next
        Exit Sub

Err:
        Dim rtn As String = "The error occur with " + control.GetType.ToString + " within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub

End Class
