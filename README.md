# MS Outlook to System Email Processor (Powershell Based)
Free to Use and Licensed under GNU GPLv3

This project accomplishs the following functions
1.  Configures Outlook with a Parent Folder and Two SubFolders
    (Processed, UnProcessed)
1a. Currently Set as JVET Inbox, Processed, and UnProcessed
2.  Scans the Parent Folder Continuously when a email is received
    it will check for specified text, and adds Text before and
    text after then sends the new email message to the defined
    From, To, SMTP Server
2a. *Recommend the emails received are in Plain-Text Only
2b. Text-Before is defined in $header1 or $header2
2c. Text-After is defined in $footer1 or footer2
3.  You can either input these settings in the command-line or set in
    the user_settings.json the following:
    SendFrom:
    SendTo:
    SMTP:
    Encryption:
    Interval:
    RI:
3a. These Settings can be taken from the command-line and saved to
    user_settings.json
4.  Performs regular clean on Parent Folder and deletes any messages
    older than 2 days   
