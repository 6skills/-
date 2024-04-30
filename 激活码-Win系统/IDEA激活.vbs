Set oShell = CreateObject("WScript.Shell")
Set oEnv = oShell.Environment("USER")
Set oFS = CreateObject("Scripting.FileSystemObject")

software = "IntelliJIdea"
javaagentFile = "idea64.exe.vmoptions"
javaagentkey = "idea.key"

Dim scriptPath, systemUser
scriptPath = oFS.GetFile(Wscript.ScriptFullName).ParentFolder.Path
systemUser = oShell.ExpandEnvironmentStrings("%APPDATA%")

jetPath = systemUser & "\JetBrains\"
if oFS.folderExists(jetPath) Then
    
Else
    MsgBox "Please launch your " & software & " first!"
    wscript.quit
End If

jetRiderCrackPath = systemUser & "\" & software

if oFS.folderExists(jetRiderCrackPath) Then
    oFS.deleteFolder jetRiderCrackPath, True 
End If

oFS.createFolder jetRiderCrackPath

oFS.createFolder jetRiderCrackPath & "\config\"
scriptPathConfig = scriptPath & "\config\"
if oFS.folderExists(scriptPathConfig) Then
    oFS.CopyFile scriptPathConfig & "*", jetRiderCrackPath & "\config\"
Else
    MsgBox "Pls unzip your crack package at all!"
    wscript.quit
End If

oFS.createFolder jetRiderCrackPath & "\plugins\"
scriptPathPlugins = scriptPath & "\plugins\"
if oFS.folderExists(scriptPathPlugins) Then
    oFS.CopyFile scriptPathPlugins & "*", jetRiderCrackPath & "\plugins\"
Else
    MsgBox "Pls unzip your crack package at all!"
    wscript.quit
End If

oFS.CopyFile scriptPath & "\active-agt.jar", jetRiderCrackPath & "\"

Exist = 0
Set allVersions = oFS.GetFolder(jetPath)
Set folders = allVersions.SubFolders
For Each folder in folders
    pos = InStr(folder.name, software) 
    If pos <> 0 Then
        Exist = pos
        oFS.CopyFile scriptPath & "\" & javaagentFile, folder.path & "\", True
        oFS.CopyFile scriptPath & "\" & javaagentkey, folder.path & "\", True
        versionPath = folder.path & "\" & javaagentFile
        ProcessVmOptions versionPath
    End If    
Next

Sub ProcessVmOptions(ByVal file)
    Dim sLine, sNewContent, bMatch
    Set oFile = oFS.OpenTextFile(file, 1, 0)

    sNewContent = ""
    Do Until oFile.AtEndOfStream
        sLine = oFile.ReadLine
        bMatch = re.Test(sLine)
        If Not bMatch Then
            sNewContent = sNewContent & sLine & vbLf
        End If
    Loop
    oFile.Close

    sNewContent = sNewContent & "--add-opens=java.base/jdk.internal.org.objectweb.asm=ALL-UNNAMED" & vbcrlf & "--add-opens=java.base/jdk.internal.org.objectweb.asm.tree=ALL-UNNAMED" & vbcrlf & "-javaagent:" & jetRiderCrackPath & "\active-agt.jar"
    Set oFile = oFS.OpenTextFile(file, 2, 0)
    oFile.Write sNewContent
    oFile.Close
End Sub

If Exist <> 0 Then
    MsgBox "Success!!! Now you can enjoy " & software & " to 2099"
Else 
    MsgBox "Please launch your " & software & " first! Then execute vbscript!!!"
End If


    


