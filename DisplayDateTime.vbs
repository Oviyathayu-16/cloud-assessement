DisplayDateTime.vbs

&#39; Function to get the current date and time
Function GetCurrentDateTime()
    Dim currentDateTime
    currentDateTime = Now  &#39; The Now function returns the current date and time
    GetCurrentDateTime = currentDateTime
End Function

&#39; Main program
Dim dateTime
dateTime = GetCurrentDateTime()

&#39; Display the date and time in a message box
MsgBox &quot;Current Date and Time: &quot; &amp; dateTime, vbInformation, &quot;Date and Time&quot;
