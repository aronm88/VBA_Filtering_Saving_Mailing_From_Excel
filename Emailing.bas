Attribute VB_Name = "Emailing"
Sub Mailing()
    
    
    ' _____IMPORTANT INFORMATION_____
    '
    ' This code requires activation of Outlook reference
    ' In VBA ribbon editor go to Tools/Reference/ and activate Microsoft Outlook xx Object Library
    ' where "xx" stand for Outlook version number
    ' user needs to use Outlook as e-mail client
    ' _______________________________
        
        
    ' declaring and creating an instances of a new Outlook objects
    Dim EmailApp As Outlook.Application
    Set EmailApp = New Outlook.Application
    Dim EmailItem As Outlook.MailItem
    
    ' declaring other variables types
    Dim File_to_attach As String
    Dim User As String
    
    ' sending an email
    For Each mail In Range("email_list")
        If mail = "" Then Exit For ' exit loop if cell does not contain any string, otherwise continue
        
        File_to_attach = mail.Offset(0, -1)
        Title = Range("Title_cell") & " - " & mail.Offset(0, -2)
        
        Set EmailItem = EmailApp.CreateItem(olMailItem)  'declaring and creating an instances of a new Outlook objects
        With EmailItem
            .To = mail
            .Subject = Title
            .Body = _
                Range("Welcome_cell") & _
                vbNewLine & _
                vbNewLine & _
                Range("body_cell") & _
                vbNewLine & _
                vbNewLine & _
                Range("greetings_cell")
            .Attachments.Add File_to_attach
            '.Save
            '.Display
            .Send
            
        End With
    Next mail
    
    
End Sub

