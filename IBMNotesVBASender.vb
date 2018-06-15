Attribute VB_Name = "Módulo1"
Option Compare Database

'Simple function for sending emails
Public Sub SendEmail(subject As String, body As String, emails() As Variant)
    
    Dim notesdb, notesdoc, notesrtf, notessession As Object
    Set notessession = CreateObject("Notes.Notessession")
    Set notesdb = notessession.GetDatabase("", "")
    notesdb.OpenMail
    Set notesdoc = notesdb.CreateDocument
    notesdoc.subject = subject
    notesdoc.SendTo = emails
    Set notesrtf = notesdoc.CreateRichTextItem("body")
    notesrtf.appendText body
    'Save message on sent items
    notesdoc.SaveMessageOnSend = True
    notesdoc.send False
    Set notessession = Nothing
    Set notesdb = Nothing
    
End Sub

'Another way informing recipients
Public Sub SendEmailString(subject As String, body As String, emails As String)
    Dim emailsendto() As String
    Dim counter As Integer
    Dim matriz
    Dim notesdb, notesdoc, notesrtf, notessession As Object
    Set notessession = CreateObject("Notes.Notessession")
    Set notesdb = notessession.GetDatabase("", "")
    notesdb.OpenMail
    Set notesdoc = notesdb.CreateDocument
    notesdoc.subject = subject
    
    matriz = Split(emails, ",")
    ReDim emailsendto(UBound(matriz))
    For counter = 0 To UBound(matriz)
        If InStr(emails, matriz(counter)) > 0 Then
            emailsendto(counter) = matriz(counter)
        End If
    Next
    
    notesdoc.SendTo = emailsendto
    Set notesrtf = notesdoc.CreateRichTextItem("body")
    notesrtf.appendText body
    'Save message on sent items
    notesdoc.SaveMessageOnSend = True
    notesdoc.send False
    Set notessession = Nothing
    Set notesdb = Nothing
    
End Sub

'Sending with Copy To and Blind Copy To
Public Sub SendEmailStringCC(subject As String, body As String, emails As String, Optional emailCC As String = "", Optional emailBCC As String = "")
    Dim emailsendto() As String
    Dim counter As Integer
    Dim matriz
    Dim notesdb, notesdoc, notesrtf, notessession As Object
    Set notessession = CreateObject("Notes.Notessession")
    Set notesdb = notessession.GetDatabase("", "")
    notesdb.OpenMail
    Set notesdoc = notesdb.CreateDocument
    notesdoc.subject = subject
    
    matriz = Split(emails, ",")
    ReDim emailsendto(UBound(matriz))
    For counter = 0 To UBound(matriz)
        If InStr(emails, matriz(counter)) > 0 Then
            emailsendto(counter) = matriz(counter)
        End If
    Next
    
    notesdoc.SendTo = emailsendto
    If Not emailCC = "" Then notesdoc.CopyTo = emailCC
    If Not emailBCC = "" Then notesdoc.BlindCopyTo = emailBCC
    Set notesrtf = notesdoc.CreateRichTextItem("body")
    notesrtf.appendText body
    'Save message on sent items
    notesdoc.SaveMessageOnSend = True
    notesdoc.send False
    Set notessession = Nothing
    Set notesdb = Nothing
    
End Sub

'Sending with attachment
Public Sub SendEmailStringCCAttach(subject As String, body As String, emails As String, Optional emailCC As String = "", Optional emailBCC As String = "", Optional attachment As String = "")
    Dim emailsendto() As String
    Dim counter As Integer
    Dim matriz
    Dim notesdb, notesdoc, notesrtf, notessession As Object
    Set notessession = CreateObject("Notes.Notessession")
    Set notesdb = notessession.GetDatabase("", "")
    notesdb.OpenMail
    Set notesdoc = notesdb.CreateDocument
    notesdoc.subject = subject
    
    matriz = Split(emails, ",")
    ReDim emailsendto(UBound(matriz))
    For counter = 0 To UBound(matriz)
        If InStr(emails, matriz(counter)) > 0 Then
            emailsendto(counter) = matriz(counter)
        End If
    Next
    
    notesdoc.SendTo = emailsendto
    If Not emailCC = "" Then notesdoc.CopyTo = emailCC
    If Not emailBCC = "" Then notesdoc.BlindCopyTo = emailBCC
    
    'Reading attachment file
    If Not attachment = "" Then
        Dim attachme As Object
        Dim embedobj As Object
        Set attachme = notesdoc.CreateRichTextItem("Attachment")
        Set embedobj = attachme.EmbedObject(1454, "", attachment, "Attachment")
    End If
    
    Set notesrtf = notesdoc.CreateRichTextItem("body")
    notesrtf.appendText body
    'Save message on sent items
    notesdoc.SaveMessageOnSend = True
    notesdoc.send False
    Set notessession = Nothing
    Set notesdb = Nothing
    
End Sub

'Sources:
'https://www.mrexcel.com/forum/excel-questions/684890-vba-code-lotus-notes-email-preparation-excel-worksheet.html#post3391183
'https://www.experts-exchange.com/questions/21069618/Appending-Signature-file-to-Body-of-an-Lotus-Notes-Email-Using-VBA.html
'http://www.alcs.ch/html-lotus-notes-email-including-html-signature-from-excel-with-vba.html
'http://www-01.ibm.com/support/docview.wss?uid=swg21098323
'https://stackoverflow.com/questions/686384/sending-formatted-lotus-notes-rich-text-email-from-excel-vba?rq=1
'https://stackoverflow.com/questions/41546887/excel-vba-send-html-email-using-ibm-notes
'https://stackoverflow.com/questions/42504385/vba-send-email-via-ibm-notes-add-signature
'https://www.mrexcel.com/forum/excel-questions/714387-attaching-signature-email-lotus-notes.html
'https://www.ozgrid.com/forum/forum/help-forums/excel-general/123532-insert-signature-on-lotus-notes-by-vba
'https://www.autoitscript.com/forum/topic/190060-halfway-solved-lotus-notes-show-signature-rich-text/
'http://www-10.lotus.com/ldd/nd85forum.nsf/GeneralCategory/991869a9221a0f1185257d23004d7053
'https://www.codeproject.com/Questions/838900/update-signature-from-file-by-lotusscript
'http://www-01.ibm.com/support/docview.wss?uid=swg21448083
'http://www-01.ibm.com/support/docview.wss?uid=swg21627014

'Sending everything in HTML format with signature option
Public Sub SendEmailStringHTML(subject As String, body As String, emails As String, Optional emailscc As String, Optional emailsbcc As String, Optional attachment As String, Optional signature As Boolean = False)
    Dim notessession As Object
    Dim notesdb As Object
    Dim notesdoc As Object
    Dim notesbody As Object
    Dim notesheader As Object
    Dim notesstream As Object
    'Dim notesmimefile As Object
    'Dim notesmimeheader As Object
    
    Set notessession = CreateObject("Notes.NotesSession")
    Set notesdb = notessession.GetDatabase("", "")
    Set notesstream = notessession.CreateStream
    notessession.convertMime = False 'Do not convert MIME to rich text
    notesdb.OpenMail
    Set notesdoc = notesdb.CreateDocument
    notesdoc.Form = "Memo"
    Set notesbody = notesdoc.CreateMIMEEntity
    
    'Set the subject
    Set notesheader = notesbody.CreateHeader("Subject")
    Call notesheader.SetHeaderVal(subject)
    
    'Set the recipients
    Set notesheader = notesbody.CreateHeader("To")
    Call notesheader.SetHeaderVal(emails)
    
    'Set Copy To
    If Not emailscc = "" Then
        Set notesheader = notesbody.CreateHeader("CC")
        Call notesheader.SetHeaderVal(emailscc)
    End If
    
    'Set Blind Copy To
    If Not emailsbcc = "" Then
        Set notesheader = notesbody.CreateHeader("BCC")
        Call notesheader.SetHeaderVal(emailsbcc)
    End If
    
    'Set Attachment file
    If Not anexo = "" Then
        Dim attachme As Object
        Dim embedobj As Object
        Set attachme = notesdoc.CreateRichTextItem("Attachment")
        Set embedobj = attachme.EmbedObject(1454, "", attachment, "Attachment")
    End If
    
    'If signature = True
    If signature Then
        'Read the IBM Notes signature
        Dim signaturelocation As String
        Dim objFSO As Object
        Dim textfile
        
        'Get the standard signature location
        signaturelocation = notesdb.getprofiledocument("CalendarProfile").GetItemValue("Signature")(0)
        Select Case notesdb.getprofiledocument("CalendarProfile").GetItemValue("SignatureOption")(0)
            Case 1 'Simple Text
                signaturelocation = Replace(signaturelocation, Chr(13), "<br>")
                body = body & signaturelocation
            Case 2 'HTML or image File
                If IsNull(signaturelocation) Or signaturelocation = "" Then
                    'Dont have signature
                Else
                    Dim line As String
                    Dim i As Integer
                    
                    Select Case UCase(Right(signaturelocation, 3))
                        Case "TXT"
                            body = body & "<br><br><br>"
                            Set objFSO = CreateObject("Scripting.FileSystemObject")
                            Set textfile = objFSO.OpenTextFile(signaturelocation, 1)
                            i = 0
                            Do Until textfile.AtEndOfStream
                                line = textfile.ReadLine
                                'i = i + 1
                                'MsgBox ThisLine
                                body = body & line & "<br>"
                            Loop
                            textfile.Close
                        Case "TML", "HTM"
                            Set objFSO = CreateObject("Scripting.FileSystemObject")
                            Set textfile = objFSO.OpenTextFile(signaturelocation, 1)
                            i = 0
                            Do Until textfile.AtEndOfStream
                                line = textfile.ReadLine
                                'i = i + 1
                                'MsgBox ThisLine
                                body = body & line
                            Loop
                            textfile.Close
                        Case "BMP"
                            body = body & "<br><br><br><img src=""data:image/bmp;base64," & EncodeFile(signaturelocation) & """/>"
                        Case "JPG", "PGE"
                            body = body & "<br><br><br><img src=""data:image/jpg;base64," & EncodeFile(signaturelocation) & """/>"
                        Case "PNG"
                            body = body & "<br><br><br><img src=""data:image/png;base64," & EncodeFile(signaturelocation) & """/>"
                        Case "GIF"
                            body = body & "<br><br><br><img src=""data:image/gif;base64," & EncodeFile(signaturelocation) & """/>"
                        Case Else
                            'Arquivo não reconhecido
                    End Select
                End If
            Case 3 'Rich Text
                body = body & "<br><br><br>" & Replace(notesdb.getprofiledocument("CalendarProfile").getfirstitem("Signature_Rich").Text, Chr(13), "<br>")
        End Select
    End If
    
    Call notesstream.WriteText(body)
    Call notesbody.SetContentFromText(stream, "text/HTML;charset=UTF-8", ENC_NONE) 'ENC_NONE, ENC_IDENTITY_7BIT or ENC_IDENTITY_8BIT
    Call notesstream.Close
    notesdoc.SaveMessageOnSend = True
    Call notesdoc.send(False)
    notessession.convertMime = True 'Restore conversion - very important
    'Call doc.Save(True, True)
    'Make mail editable by user
    'CreateObject("Notes.NotesUIWorkspace").EDITDOCUMENT True, doc
    'Could send it here
    Set notessession = Nothing
    Set notesdb = Nothing
    
End Sub

'Sources:
'https://stackoverflow.com/questions/41638124/vba-convert-a-binary-image-to-a-base64-encoded-string-for-a-webpage
'https://stackoverflow.com/questions/2043393/convert-image-jpg-to-base64-in-excel-vba

'Reference to Microsoft XML, v6.0 (or v3.0) required
Public Function EncodeFile(strPicPath As String) As String
    Const adTypeBinary = 1  'Binary file is encoded

    'Variables for encoding
    Dim objXML
    Dim objDocElem

    'Variable for reading binary picture
    Dim objStream

    'Open data stream from picture
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = adTypeBinary
    objStream.Open
    objStream.LoadFromFile (strPicPath)

    'Create XML Document object and root node that will contain the data
    Set objXML = CreateObject("MSXml2.DOMDocument")
    Set objDocElem = objXML.createElement("Base64Data")
    objDocElem.DataType = "bin.base64"

    'Set binary value
    objDocElem.nodeTypedValue = objStream.Read()

    'Get base64 value
    EncodeFile = objDocElem.Text

    'Clean all
    Set objXML = Nothing
    Set objDocElem = Nothing
    Set objStream = Nothing

End Function
