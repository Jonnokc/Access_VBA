Private Sub Form_BeforeUpdate(Cancel As Integer)
' perform data validation

' Will need to change Control Field to match each use
    If IsNull(Me.ClientMnemonicID) Then
        MsgBox "You must enter a New Blacklisted Code System.", vbCritical, "Data entry error..."
        DoCmd.GoToControl "BlacklistedCodeSystem"

        Cancel = True
    End If
    If Not Cancel Then
        ' passed the validation process
        If Me.NewRecord Then
            If MsgBox("Data will be saved, Are you Sure?", vbYesNo, "Confirm") = vbNo Then
                Cancel = True
            Else
                ' run code for new record before saving

            End If

        Else
            If MsgBox("Data will be modified, Are you Sure?", vbYesNo, "Confirm") = vbNo Then
                Cancel = True
            Else
                ' run code before an existing record is saved
                ' example: update date last modified

            End If
        End If
    End If
    ' if the save has been canceled or did not pass the validation , then ask to Undo changes
    If Cancel Then
        If MsgBox("Do you want to undo all changes?", vbYesNo, "Confirm") = vbYes Then
            Me.Undo
        End If

    End If
End Sub
