Public Sub CustomMailMessageRuleCreateCalendarItemHeb(Item As Outlook.MailItem)
   
   Dim objNameSpace As Outlook.NameSpace
   Dim myRequiredAttendee As Outlook.Recipient
   Dim objCopy As Outlook.MailItem
   Set objNameSpace = Application.GetNamespace("MAPI")
   Set objOwner = objNameSpace.CreateRecipient("siftingtm@gmail.com")
   Dim objRecipients As Outlook.Recipients
   Dim objRecipient As Outlook.Recipient
   Dim myAttachments As Outlook.Attachments
   
   ' to write to .ics file:
   Dim CurrentTextFile As TextStream
   Dim PathString As String
   Dim Filename As String
   Dim FullPath As String
   Dim DataString As String
   

   
   'Definitions for creating outlook calendar item
 Dim calendarItem As Object

Const olFolderCalendar = 9
Const olAppointmentItem = 1 '1 = Appointment


   
   'We need to parse the incoming message to extract information from it.
   ' Set reference to VB Script library
   ' Microsoft VBScript Regular Expressions 5.5
   
   Dim Reg1 As RegExp
   Dim Reg2 As RegExp
   Dim bookingDay As String
   Dim bookingTime As String
   Dim M1 As MatchCollection
   Dim M As Match
   Dim N1 As MatchCollection
   Dim N As Match
   Dim strSubject As String
   Dim testSubject As String
   Dim extraMarkupForBody As String
   Dim CustomerName As String
   Dim bookingCustomer As String              'name of sifting participant
   Dim bookingDate As String
   Dim bookingDateUSFormat As String
   Dim bookingNumTickets As String
   Dim bookingLang As String
   Dim bookingTicketType As String
   Dim bookingPhoneWithSpaces As String
   Dim bookingPhone As String
   Dim HebrewName As String
   Dim bookingReference As String
   
   Dim searchStringSiftingLang As String
   Dim searchStringCustomer As String
   Dim searchStringCustomerName As String
   Dim searchStringDate As String
   Dim searchStringTime As String
   Dim searchStringPhone As String
   Dim searchStringNumTickets As String
   
   'Hebrew characters in unicode - see chart here: https://www.ssec.wisc.edu/~tomw/java/unicode.html#x0590
   
   searchStringSiftingLang = ChrW$(1488) & ChrW$(1504) & ChrW$(1490) & ChrW$(1500) & ChrW$(1497) & ChrW$(1514) ' 'àðâìéú
   searchStringCustomer = ChrW$(1492) & ChrW$(1502) & ChrW$(1494) & ChrW$(1502) & ChrW$(1497) & ChrW$(1503) ' 'äîæîéï
   searchStringCustomerName = ChrW$(1513) & ChrW$(1501) ' 'ùí
   searchStringPhone = ChrW$(1496) & ChrW$(1500) & ChrW$(1508) & ChrW$(1493) & ChrW$(1503)  ' 'èìôåï
   searchStringDate = ChrW$(1514) & ChrW$(1488) & ChrW$(1512) & ChrW$(1497) & ChrW$(1498)    ' 'úàøéê
   searchStringTime = ChrW$(1493) & ChrW$(1513) & ChrW$(1506) & ChrW$(1492)    ' 'åùòä
   searchStringNumTickets = ChrW$(1499) & ChrW$(1502) & ChrW$(1493) & ChrW$(1514)    ' 'ëîåú
   searchStringReferenceNum = ChrW$(1492) & ChrW$(1494) & ChrW$(1502) & ChrW$(1504) & ChrW$(1514) & ChrW$(1499) & ChrW$(1501) & ChrW$(32) & ChrW$(1502) & ChrW$(1505) & ChrW$(1508) & ChrW$(1512)   'äæîðúëí îñôø'
   
   Set Reg1 = New RegExp
   Set Reg2 = New RegExp
   ' \s* = invisible spaces
   ' \d* = match digits
   ' \w* = match alphanumeric
   ' \S  Matches non-whitespace
   
    With Reg1
        '.Pattern = "Dear (\w*\s*\S*) ,"
        'Note: The . (dot or period) character in a regular expression is a wildcard character that matches any character except \n
         .Pattern = "Dear (.+) ,"
        .Global = True
    End With
    If Reg1.Test(Item.Body) Then
    'Confirmation email is in English
    'Extract customer name:
        Set M1 = Reg1.Execute(Item.Body)
        For Each M In M1
            ' M.SubMatches(1) is the (\w*) in the pattern
            ' use M.SubMatches(2) for the second one if you have two (\w*)
             bookingCustomer = M.SubMatches(0)
        Next
    
    'Now checking which sifting language the booking ordered:
          With Reg1
            .Pattern = " Eng "
            .Global = True
        End With
        If Reg1.Test(Item.Body) Then
            bookingLang = " ENG "
        Else
            bookingLang = " HEB "
        End If
        
         CustomerName = bookingCustomer
        'Testing for number of tickets
         With Reg1
            .Pattern = "Ticket Quantity\s*(\d*)\s*Full Name"
            .Global = True
        End With
        If Reg1.Test(Item.Body) Then
    
            Set M1 = Reg1.Execute(Item.Body)
            For Each M In M1
                Debug.Print M.SubMatches(0)
                bookingNumTickets = M.SubMatches(0)
            Next
        End If
    
        'Looking for phone number
        'note: Empty lines are not allowed in some versions of .ics usage (Google calendar)
        With Reg1
            .Pattern = "Phone\s*(\d*)\s*\S"
            .Global = True
        End With
        If Reg1.Test(Item.Body) Then
    
            Set M1 = Reg1.Execute(Item.Body)
            For Each M In M1
                Debug.Print M.SubMatches(0)
                bookingPhone = M.SubMatches(0)
            Next
        End If
    
        'Looking for time & date of booking
        With Reg1
            .Pattern = "Date and Time\s*\s*\n*\r*\s*(\S* \S*)\s*\n*\r*Ticket Quantity"
            .Global = True
        End With
        If Reg1.Test(Item.Body) Then
        
            Set M1 = Reg1.Execute(Item.Body)
            For Each M In M1
                Debug.Print M.SubMatches(0)
                bookingDate = M.SubMatches(0)
            Next
        End If
    
        'MsgBox "Direct reservation Date: " & bookingDate 'successful
        
        'Looking for booking reference num
        With Reg1
            .Pattern = "Reservation Number\s*\s*\n*\r*\s*(\d*)\s*\n*\r*Date and Time"
            .Global = True
        End With
        If Reg1.Test(Item.Body) Then
        
            Set M1 = Reg1.Execute(Item.Body)
            For Each M In M1
                Debug.Print M.SubMatches(0)
                bookingReference = M.SubMatches(0)
            Next
        End If
    
        'MsgBox "Booking ref num: " & bookingReference 'successful

    'End of text extraction from confirmation email in English
    Else
        
    'Confirmation email is in Hebrew
    
        'Now checking which sifting language the booking ordered:
        'If oredered via the English site, the reservation message will be in English, but they may have booked either language
        
         With Reg1
            .Pattern = searchStringSiftingLang      'Looks for אנגלית
            .Global = True
        End With
        If Reg1.Test(Item.Body) Then
            bookingLang = " ENG "
        Else
            bookingLang = " HEB "
        End If
        
       'Searching for customer name:
        With Reg1
          ' the ChrW$ Function Returns the Unicode character that corresponds to the specified character code.
            'Looking for characters after "שם המזמין" up until new line
            'see https://stackoverflow.com/questions/44943450/vba-regex-matching-fraction-characters-e-g-1-8-3-8-in-string
            'dot matches every character except newline
        'ChrW$(32) represents a single space
     .Pattern = searchStringCustomerName & ChrW$(32) & searchStringCustomer & "(.+\s*)"   'Looks for words after שם המזמין
     
         'MsgBox "Testing for: " & searchStringCustomerName & ChrW$(32) & searchStringCustomer
            .Global = True
        End With
        
        If Reg1.Test(Item.Body) Then
            Set M1 = Reg1.Execute(Item.Body)
            For Each M In M1
                bookingCustomer = M.SubMatches(0) 'This will give the name and telephone together
                'MsgBox "From Heb site, Name of customer: " & bookingCustomer
            Next

        Else
        Exit Sub
        End If
        
    'Searching for telephone number
    ' ChrW$(32) = space
    
        With Reg1
          .Pattern = searchStringPhone & ChrW$(32) & searchStringCustomer & "\s*(\d*)\s*\S"   'Looks for numerical characters only after טלפון המזמין
        'MsgBox "Testing for: " & searchStringPhone & ChrW$(32) & searchStringCustomer
            .Global = True
        End With
        
        If Reg1.Test(Item.Body) Then
            Set M1 = Reg1.Execute(Item.Body)
            For Each M In M1
                bookingPhoneWithSpaces = M.SubMatches(0) 'This will give the telephone together with some whitespace
                bookingPhone = Trim(bookingPhoneWithSpaces) 'This is meant to strip off the whitespace
                'MsgBox "Telephone of customer: " & bookingPhone
            Next

        Else
        Exit Sub
        End If
        
        'searchStringDate
        With Reg1
          .Pattern = searchStringDate & ChrW$(32) & searchStringTime & "(.+\s*)"   'Looks for words after תאריך ושעה
     
        'MsgBox "Testing for: " & searchStringDate & ChrW$(32) & searchStringTime
            .Global = True
        End With
        
        If Reg1.Test(Item.Body) Then
            Set M1 = Reg1.Execute(Item.Body)
            For Each M In M1
                bookingDate = M.SubMatches(0) 'This will give the name and telephone together
                'MsgBox "Date: " & bookingDate
            Next

        Else
        Exit Sub
        End If
         
         'Search for number of tickets
         With Reg1
          .Pattern = searchStringNumTickets & "(\d*)"   'Looks for words after כמות
     
        'MsgBox "Testing for: " & searchStringNumTickets
            .Global = True
        End With
        
        If Reg1.Test(Item.Body) Then
            Set M1 = Reg1.Execute(Item.Body)
            For Each M In M1
                bookingNumTickets = M.SubMatches(0) 
                'MsgBox "Number of tickets: " & bookingNumTickets
                Exit For 'don't move to the next instance of כמות
            Next

        Else
        Exit Sub
        End If

        With Reg1
          .Pattern = searchStringReferenceNum & "\s*(\d*)\s*\S"   'Looks for words after הזמנתכם מספר
     
        MsgBox "Testing for: " & searchStringReferenceNum
            .Global = True
        End With
        
         If Reg1.Test(Item.Body) Then
            Set M1 = Reg1.Execute(Item.Body)
            For Each M In M1
                bookingReference = M.SubMatches(0) 'This will give the booking reference number
                MsgBox "Booking reference: " & bookingReference
                Exit For 'don't move to the next instance of הזמנתכם מספר
            Next

        Else
        Exit Sub
        End If
        
         
    End If
     
  'MsgBox "Customer name: " & bookingCustomer
     
    bookingCustomer = bookingCustomer & " " & bookingLang & " " & bookingNumTickets & " " & bookingPhone
    
'MsgBox "Customer name: " & bookingCustomer & " Date: " & bookingDate
 

    
      If IsDate(bookingDate) Then
        If (Hour(bookingDate) = 12) Then
            If (Minute(bookingDate) = 0) Then
                bookingTime = Hour(bookingDate) & ":" & Minute(bookingDate) & "0 PM"
            Else
                bookingTime = Hour(bookingDate) & ":" & Minute(bookingDate) & " PM"
            End If
        ElseIf (Hour(bookingDate) > 12) Then
            If (Minute(bookingDate) = 0) Then
                bookingTime = (Hour(bookingDate) - 12) & ":" & Minute(bookingDate) & "0 PM"
            Else
                bookingTime = (Hour(bookingDate) - 12) & ":" & Minute(bookingDate) & " PM"
            End If
        Else
            If (Minute(bookingDate) = 0) Then
                bookingTime = Hour(bookingDate) & ":" & Minute(bookingDate) & "0 AM"
            Else
                bookingTime = Hour(bookingDate) & ":" & Minute(bookingDate) & " AM"
            End If
        End If
    End If
      
     bookingDateUSFormat = "# " & Month(bookingDate) & "/" & Day(bookingDate) & "/" & Year(bookingDate) & " " & bookingTime & " #"
    'MsgBox "US date format " & bookingDateUSFormat
   With Reg1
        .Pattern = "Currency(\s*\s*\n*\r*\s*(\S* \S*)\s*\n*\r*)Total"
        .Global = True
    End With
    If Reg1.Test(Item.Body) Then
    
        Set M1 = Reg1.Execute(Item.Body)
        For Each M In M1
            Debug.Print M.SubMatches(0)
            bookingTicketType = M.SubMatches(0)
        Next
    End If
    
   'MsgBox "creating outlook calendar event " 'successful
   
  'creating outlook calendar event
  Set calendarItem = Application.CreateItem(olAppointmentItem)
  calendarItem.MeetingStatus = olMeeting
 
 calendarItem.Subject = bookingCustomer  'title of appt - name of booker, lang, num tickets
 calendarItem.Location = "Mitzpeh HaMasuot"
 
 calendarItem.Start = bookingDate
  
 'Request from sifting staff for booking duration to be 30 mins
 'In google calendar, set default event length in Event Settings
 calendarItem.Duration = 30
 Set myRequiredAttendee = calendarItem.Recipients.Add("siftingtm@gmail.com")
  
 myRequiredAttendee.Type = olRequired
 'calendarItem.Display
 '*****************************
    
 'MsgBox "wrtingn info to text file"
    ' Create the ics file
    Filename = bookingReference + ".ics"
    PathString = "C:\temp\"
    FullPath = PathString + Filename

        'if you need to check that file does not already exist, this is the code you need:
    MsgBox "trying file " & FullPath
    If Not Dir(FullPath) = "" Then
        MsgBox "file exists"
        GoTo FileExists
        'File with this name already exists, exit sub
    End If


    
    ' Creation of .ics file - this will be sent as an attachment
    ' and automatically imported by gmail into google calendar.
       'will use fst (stream-type file with utf-8) instead of fso  for the file attachment.
   Dim fs, fsT As Object
 
  '**************************************
 ' different ways to create text file:
 ' *****
 ' creating a FileSystemObject and using CreateTextFile and WriteLine/Write:
 ' ***
 '  Set fs = CreateObject("Scripting.FileSystemObject")
 '  Set a = fs.CreateTextFile("c:\testfile.txt", True)
 '  a.Write ("This is a test.")
 '  a.WriteLine ("This is another test.")  'writeline adds a newline character to the end of the string
 '  a.Close
 ' *****
 ' creating a ADODB.Stream (specifying type 2 which is for text/string data), and using writetext:
 ' ***
 ' Set fsT = CreateObject("ADODB.Stream")
 ' fsT.Type = 2
 ' fsT.Charset = "utf-8"
 ' fsT.Open
 ' fsT.writetext "test text", 1
 ' fsT.SaveToFile FullPath, 2   '2: to allow overwrite
 ' fsT.Close
 '
 ' *****
 
  'Using the second method:
  'Create Stream object
'    Set fsT = CreateObject("ADODB.Stream")
'    'Specify stream type - we want To save text/string data.
'    fsT.Type = 2
'    'Specify charset For the source text data.
'    fsT.Charset = "utf-8"
'    'Open the stream And write binary data To the object
'    fsT.Open
'    MsgBox "opened file for writing: " & FullPath
'    ' Write header information
'    'syntax: objStream.WriteText data,opt
'    'opt: optional paramater:  1   Writes the specified text and a line separator to a Stream object.
'    DataString = "BEGIN:VCALENDAR"
'    fsT.writetext DataString, 1
'    DataString = "BEGIN:VEVENT"
'    fsT.writetext DataString, 1
'    DataString = "SUMMARY:" + bookingCustomer
'    fsT.writetext DataString, 1
'    DataString = "LOCATION:" + "Mitzpeh HaMasuot"
'    fsT.writetext DataString, 1
'    DataString = "DTSTART;VALUE=DATE:" + Format(bookingDate, "yyyymmdd") + "T" + Format(bookingDate, "hhmmss")
'    fsT.writetext DataString, 1
'    DataString = "DESCRIPTION:" + bookingPhone
'    fsT.writetext DataString, 1
'    DataString = "END:VEVENT"
'    fsT.writetext DataString, 1
'    DataString = "END:VCALENDAR"
'    fsT.writetext DataString, 1
'       'Save (binary) data To disk
'       MsgBox "Saving ics file" & FullPath
'    fsT.SaveToFile FullPath, 2 '2 = overwrite
'    fsT.Close

  'Using the first method:
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set fsT = fs.CreateTextFile(FullPath, True)
    fsT.WriteLine ("BEGIN:VCALENDAR")
    fsT.WriteLine ("BEGIN:VEVENT")
    fsT.WriteLine ("SUMMARY:" + bookingCustomer)
    fsT.WriteLine ("LOCATION:" + "Mitzpeh HaMasuot")
    fsT.WriteLine ("DTSTART;VALUE=DATE:" + Format(bookingDate, "yyyymmdd") + "T" + Format(bookingDate, "hhmmss"))
    fsT.WriteLine ("DESCRIPTION:" + bookingPhone)
    fsT.WriteLine ("END:VEVENT")
    fsT.WriteLine ("END:VCALENDAR")
    fsT.Close
    MsgBox "closed file"

 Set objRecipients = Item.Recipients

    For Each objRecipient In objRecipients
        strRecipientAddress = objRecipient.Address
        'Find & Delete the specific recipient
        'Change the source recipient address as per your own case
        If strRecipientAddress = "booking@tmsifting.org" Then
          objRecipient.Delete
          Exit For
        End If
    Next
 
 'MsgBox "Sending message"
 
    'put in siftingtm@gmail.com as the new To: address
   Item.Recipients.Add "siftingtm@gmail.com"
  
   'Item.SentOnBehalfOfName = "siftingtm@gmail.com"
   Set myAttachments = Item.Attachments
   
   Item.Attachments.Add FullPath
   'put the message in the outbox
   '*********************
    'temporarily commenting this out since don't want now to add to diary
   Item.Send
   
 ' calendarItem.Display
 'Sending outlook calendar item"
 'calendarItem.Send
 
Exit Sub

ErrorHandl:
   Debug.Print "Error number: " & Err.Number
    Err.Clear
FileExists:
    Debug.Print "File already exists"
    MsgBox "File already exists"
    Err.Clear
End Sub