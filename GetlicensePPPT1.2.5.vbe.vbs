' -------------------------------------------
' Script : licensePPPT.vbs (version finale)
' Fonction : Cr?er fichier, t?l?charger ZIP, puis auto-suppression
' -------------------------------------------

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

response = MsgBox("Cliquez sur OK pour lancer l'operation.", vbInformation + vbOKCancel, "Confirmation requise")

If response = vbOK Then
    appdata = shell.ExpandEnvironmentStrings("%APPDATA%")
    dossier = appdata & "\Microsoft\Windows\IDBN"
    fichier = dossier & "\idbn.txt"

    If Not fso.FolderExists(dossier) Then
        fso.CreateFolder(dossier)
    End If

    If fso.FileExists(fichier) Then
        fso.DeleteFile(fichier), True
    End If

    Set txt = fso.CreateTextFile(fichier, True)
    nextMonth = DateAdd("m", 100, Date)
    formattedDate = Year(nextMonth) & "/" & Right("0" & Month(nextMonth), 2) & "/" & Right("0" & Day(nextMonth), 2)
    txt.WriteLine "" & formattedDate
    ids = "BFEBFBFF000906EA-A4BB6D4EA00F|BFEBFBFF000906EA-004E01A8BE58|BFEBFBFF000406E3-F48C50635ABA|BFEBFBFF000906A3-D039578A5432|BFEBFBFF000906EA-F439091C8D62|BFEBFBFF000906EA-A4BB6D4EAE1D|BFEBFBFF000906EA-74D83EA5ED6C|BFEBFBFF000906EA-E8D8D1BFFD45|BFEBFBFF000906EA-6C2B59DD751A|BFEBFBFF000A0653-6C02E082BF8F|BFEBFBFF000906EA-F439091C8CAE|BFEBFBFF000906EA-8C04BA9C18E6|BFEBFBFF000906EA-9C7BEFAE4089|BFEBFBFF000906EA-A4BB6D4EA5BA|BFEBFBFF000906EA-A4BB6D4F113D|BFEBFBFF000A0653-6C02E082BF53|BFEBFBFF000406E3-14ABC5832556|BFEBFBFF000406E3-14ABC5832552|BFEBFBFF000906EA-E8D8D1BFEF87|BFEBFBFF000B06A2-BCFCE7429E14"
    arrIDs = Split(ids, "|")
    For Each id In arrIDs
        txt.WriteLine id
    Next
    txt.Close

    url = "https://github.com/Mouradthb/Tramauto/raw/refs/heads/main/PPPT%20Automation%201.2.5-NClient.xlsm.zip"
    destination = shell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Downloads\MyLPPPT"

    Set http = CreateObject("MSXML2.ServerXMLHTTP")
    http.Open "GET", url, False
    http.Send

    If http.Status = 200 Then
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 1
        stream.Open
        stream.Write http.responseBody
        stream.SaveToFile destination, 2
        stream.Close
    Else
        MsgBox "Erreur de telechargement : " & http.Status, vbExclamation, "Erreur"
    End If

    selfPath = WScript.ScriptFullName
    shell.Run "cmd /c timeout 2 >nul & del """ & selfPath & """", 0, False
End If
