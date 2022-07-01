On Error GoTo Handler

Dim str_input As String
Dim str_output As String
Dim errnum as Integer
Dim str_manufacturer As String
Dim str_model As String
Dim str_fixed As String
Dim str_version As String

Open "COM3:9600,N,8,1,LF" For Random As #1
If errnum = 0 Then
   Print "[INFO] COM opened."
Else
   Print "[ERROR] COM failed to open.  ErrNum="; errnum
End If

Print "[INFO] Querying device..."
Print #1, "*IDN?\r\n"
Input #1, str_manufacturer
Input #1, str_model
Input #1, str_fixed
Input #1, str_version

Close #1
If errnum = 0 Then
   Print "[INFO] COM closed."
Else
   Print "[ERROR] COM failed to close.  ErrNum="; errnum
End If
Print ""
Print "Manufacturer: "+str_manufacturer+", Model: "+str_model+", Fixed: "+str_fixed+", Ver: "+str_version
Print ""
Print "Press ENTER to exit"
Input str_input
End

Handler:
errnum = Err
Print "Error:"; errnum
