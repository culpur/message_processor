# MS Outlook to System Email Processor (Powershell Based)
Free to Use and Licensed under GNU GPLv3

This project accomplishs the following functions
1.  Configures Outlook with a Parent Folder and Two SubFolders
    (Processed, UnProcessed)
2. Currently Set as JVET Inbox, Processed, and UnProcessed
3.  Scans the Parent Folder Continuously when a email is received
    it will check for specified text, and adds Text before and
    text after then sends the new email message to the defined
    From, To, SMTP Server
4. *Recommend the emails received are in Plain-Text Only
5. Text-Before is defined in $header1 or $header2
6. Text-After is defined in $footer1 or footer2
7.  You will be asked to input settings in the command-line or 
    set alternatively you can set these settings in the user_settings.json.
    SendFrom:
    SendTo:
    SMTP:
    Encryption:
    Interval:
    RI:
8. These Settings can be taken from the command-line and saved to
    user_settings.json
9.  Performs regular clean on Parent Folder and deletes any messages
    older than 2 days   
