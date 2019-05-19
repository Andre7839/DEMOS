Imports System.Windows.Forms
Public Class Interfaces
    Function assign_rough(ByVal dg_db As DataGridView, ByVal dg_out As DataGridView, ByVal col_out As Integer, ByVal col_db As Integer) As Single()
        Try


            Dim i, j As Integer
            Dim values As Array = Array.CreateInstance(GetType(Single), dg_out.Rows.Count + 1)

            For i = 1 To dg_out.Rows.Count
                For j = 1 To dg_db.Rows.Count
                    If dg_db.Item(col_db - 1, j - 1).Value = dg_out.Item(col_out - 1, i - 1).Value Then
                        values(i) = dg_db.Item(col_db, j - 1).Value

                    End If
                Next
            Next
            Return values
        Catch ex As Exception

        End Try


    End Function

   

    Function copy_all_row(ByVal DG As DataGridView, ByVal begin_col As Integer, ByVal end_col As Integer) As DataGridView
        DG.EndEdit()
        Dim ii, i, j As Integer
        For ii = 0 To DG.Rows.Count - 1
            If DG.Rows(ii).Selected = True Then

                For j = ii + 1 To DG.Rows.Count - 1
                    For i = begin_col To end_col
                        DG.Item(i, j).Value = DG.Item(i, ii).Value
                    Next
                Next


            End If
        Next
        Return DG
    End Function
    Function copy_row_step(ByVal DG As DataGridView, ByVal begin_col As Integer, ByVal end_col As Integer) As DataGridView
        DG.EndEdit()
        Try


            Dim ii, i, f As Integer

            For ii = 0 To DG.Rows.Count - 1

                If DG.Rows(ii).Selected = True Then
                    f = ii
                    If ii = DG.Rows.Count - 1 Then
                        Exit Function
                    End If
                    For i = begin_col To end_col
                        DG.Item(i, ii + 1).Value = DG.Item(i, ii).Value
                    Next
                End If
            Next

            DG.Rows(f).Selected = False
            DG.Rows(f + 1).Selected = True
            Return DG
        Catch ex As Exception

        End Try
    End Function

    Function copy_column(ByVal DG As DataGridView, ByVal start_column As Integer) As DataGridView
        DG.EndEdit()
        Try
            Dim ii, i, j As Integer
            For ii = start_column To DG.Columns.Count
                For i = 1 To DG.Rows.Count
                    If DG.Item(ii - 1, i - 1).Selected = True Then
                        For j = i To DG.Rows.Count
                            DG.Item(ii - 1, j - 1).Value = DG.Item(ii - 1, i - 1).Value
                        Next
                    End If
                Next
            Next
            Return DG
        Catch ex As Exception

        End Try

    End Function

    Public Sub TextBox_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        '---if textbox is empty and user pressed a decimal char---
        If CType(sender, TextBox).Text = String.Empty And e.KeyChar = Chr(46) Then
            e.Handled = True
            Return
        End If
        '---if textbox is empty and user pressed negative sign
        If CType(sender, TextBox).Text = String.Empty And e.KeyChar = "-" Then
            e.Handled = False
            Return
        End If

        '---if textbox already has a decimal point---
        If CType(sender, TextBox).Text.Contains(Chr(46)) And e.KeyChar = Chr(46) Then
            e.Handled = True
            Return
        End If
        '---if the key pressed is not a valid decimal number---
        If (Not (Char.IsDigit(e.KeyChar) Or Char.IsControl(e.KeyChar) Or (e.KeyChar = Chr(46)))) Then
            e.Handled = True
        End If
    End Sub


    Sub double_keypress(ByVal str As String, ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        'this sub is used to assign only decimal numbers

        Dim i As Integer = str.IndexOf(".")

        If Char.IsNumber(e.KeyChar) = False Then
            If e.KeyChar = CChar(ChrW(Keys.Back)) Then
                e.Handled = False
                Exit Sub
            End If
            If e.KeyChar = CChar(".") Then
                If Not i = -1 Then
                    e.Handled = True
                    Exit Sub
                End If
                e.Handled = False
                Exit Sub
            End If
            e.Handled = True
        End If

    End Sub

    Sub integer_keypress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        'this sub is used to assign only integer numbers

        If Char.IsNumber(e.KeyChar) = False Then
            If e.KeyChar = CChar(ChrW(Keys.Back)) Then
                e.Handled = False
                Exit Sub
            End If
            e.Handled = True
        End If
    End Sub
End Class
