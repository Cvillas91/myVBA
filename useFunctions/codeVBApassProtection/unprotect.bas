Attribute VB_Name = "unprotect"

'RUN THIS SUBROUTINE TO UNPROTECT ALL MACRO PASSWORDS

Sub unprotected()
    If Hook Then
        MsgBox "VBA Project is unprotected!", vbInformation, "*****"
    End If
End Sub
