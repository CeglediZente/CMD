Sample = msgBox("Option codes 0+64." & vbCrLf & vbCrLf & "This is an information situation with OK only",0+64, "Information.")
    Select Case Sample
        Case 1
            wScript.Echo "You acknowledged."
    End Select

Sample = msgBox("Option codes 1+48." & vbCrLf & vbCrLf & "This is a warning situation with OK and Cancel",1+48, "WARNING!")
    Select Case Sample
        Case 1
            wScript.Echo "You chose OK."
        Case 2
            wScript.Echo "You chose Cancel."
    End Select

Sample = msgBox("Option codes 2+16." & vbCrLf & vbCrLf & "This is a critical situation with Abort, Retry, and Ignore",2+16, "CRITICAL!")
    Select Case Sample
        Case 3
            wScript.Echo "You chose Abort."
        Case 4
            wScript.Echo "You chose Retry."
    Case 5
        wScript.Echo "You chose Ignore."
    End Select

Sample = msgBox("Option codes 0+16." & vbCrLf & vbCrLf & "This is a critical error situation with OK only",0+16, "CRITICAL ERROR!")
    Select Case Sample
        Case 1
            wScript.Echo "You acknowledged."
    End Select

Sample = msgBox("Option codes 3+32." & vbCrLf & vbCrLf & "This is a question situation with Yes, No, or Cancel",3+32, "Question?")
    Select Case Sample
        Case 6
            wScript.Echo "You said Yes."
        Case 7
            wScript.Echo "You said No."
        Case 2
            wScript.Echo "You clicked Cancel."
    End Select

Sample = msgBox("Option codes 4+32." & vbCrLf & vbCrLf & "This is a question situation with Yes or No only",4+32, "Question?")
    Select Case Sample
        Case 6
            wScript.Echo "You said Yes."
        Case 7
            wScript.Echo "You said No."
    End Select

Sample = msgBox("Option codes 5+16." & vbCrLf & vbCrLf & "This is a critical error situation with Retry or Cancel",5+16, "CRITICAL ERROR!")
    Select Case Sample
        Case 4
            wScript.Echo "You clicked Retry."
        Case 2
            wScript.Echo "You clicked Cancel."
    End Select