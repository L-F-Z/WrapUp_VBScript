' # Wrap Script for Computer Architecture Lab Course
' # Please Put It at the Same Folder with 'mycpu_verify' & 'cpu132_gettrace'
' $Func    Use Windows Build-in Zip Function to Backup or FinalWrap
' $Coder   LFZ
' $Ver     0.9(20181008)
' $Src     https://github.com/L-F-Z/WrapUp_VBScript

'|-lab#_2016K80099XXXXX/
'| |--lab#_2016K80099XXXX.pdf/
'| |-mycpu_verify/
'| | |--rtl/
'| | | |--soc_lite_top.v
'| | | |--myCPU /
'| | | |--CONFREG/
'| | | |--BRIDGE/
'| | | |--xilinx_ip/
'| | |--testbench/
'| | | |--mycpu_tb.v
'| | |--run_vivado/
'| | | |--soc_lite.xdc
'| | | |--mycpu/
'| | | | |--mycpu.xpr
'| | | | |--mycpu.bit

Set fso = CreateObject("Scripting.FileSystemObject")
Dim choice
choice = msgbox("Yes = Backup    No = FinalWrap    Cancel = Exit", vbYesNoCancel, "Wrap Script by LFZ")
select case choice
    case 6
        Backup()
    case 7
        FinalWrap()
    case Else
        WScript.Quit
end select

'--------------------------------------Functions-----------------------------------------
Function FinalWrap()
    Dim ID, LAB, PRJ
    ID = InputBox("Please Enter Your ID: ", "Get ID", "2016K80099XXXXX")
    LAB = InputBox("Please Enter Lab ID: ", "Get Lab ID", "lab1_")
    PRJ = InputBox("Please Enter Project Name: ", "Get Project Name", "mycpu_prj1")
    'IF YOU DO NOT WANT TO ENTER THEM EVERY TIME, JUST USE THE FOLLOWING INSTRUCTIONS
    'ID = "2016K80099XXXXX"
    'LAB = "lab1_"
    'PRJ = "mycpu_prj1"
    
    Dim DIR,ReportFile, WrapDir, FileName
    DIR = GetCurrentFolder()
    CreateFolder DIR&"\ZipFolderTmp"
    CreateFolder DIR&"\ZipFolderTmp\"&LAB&ID
    
    msgbox "Please Select Your Project Report"
    ReportFile = SelectFile()
    Filename   = GetFileName(ReportFile)
    WrapDir    = DIR&"\ZipFolderTmp\"&LAB&ID
    CopyFile     ReportFile, WrapDir
    FileRename   WrapDir&"\"&FileName, LAB&ID&".pdf"
    
    CreateFolder DIR&"\ZipFolderTmp\"&LAB&ID&"\mycpu_verify"
    CreateFolder DIR&"\ZipFolderTmp\"&LAB&ID&"\mycpu_verify\rtl"
    CopyFolder   DIR&"\mycpu_verify\rtl", DIR&"\ZipFolderTmp\"&LAB&ID&"\mycpu_verify\rtl"
    CreateFolder DIR&"\ZipFolderTmp\"&LAB&ID&"\mycpu_verify\testbench"
    CopyFolder   DIR&"\mycpu_verify\testbench", DIR&"\ZipFolderTmp\"&LAB&ID&"\mycpu_verify\testbench"
    CreateFolder DIR&"\ZipFolderTmp\"&LAB&ID&"\mycpu_verify\run_vivado"
    CopyFile     DIR&"\mycpu_verify\run_vivado\soc_lite.xdc", DIR&"\ZipFolderTmp\"&LAB&ID&"\mycpu_verify\run_vivado"
    CreateFolder DIR&"\ZipFolderTmp\"&LAB&ID&"\mycpu_verify\run_vivado\"&PRJ
    CopyFile     DIR&"\mycpu_verify\run_vivado\"&PRJ&"\"&PRJ&".runs\impl_1\soc_lite_top.bit", DIR&"\ZipFolderTmp\"&LAB&ID&"\mycpu_verify\run_vivado\"&PRJ
    CopyFile     DIR&"\mycpu_verify\run_vivado\"&PRJ&"\"&PRJ&".xpr", DIR&"\ZipFolderTmp\"&LAB&ID&"\mycpu_verify\run_vivado\"&PRJ
    
    CreateZip    DIR&"\ZipFolderTmp", DIR&"\"&LAB&ID&".zip"
    wScript.Sleep 10000
    DeleteFolder DIR & "\ZipFolderTmp"
    WScript.Quit
End Function

Function Backup()
    Dim LAB, PRJ
    LAB = InputBox("Please Enter Lab ID: ", "Get Lab ID", "lab1-")
    PRJ = InputBox("Please Enter Project Name: ", "Get Project Name", "mycpu_prj1")
    'IF YOU DO NOT WANT TO ENTER THEM EVERY TIME, JUST USE THE FOLLOWING INSTRUCTIONS
    'LAB = "lab1_"
    'PRJ = "mycpu_prj1"
    
    Dim DIR,strFolder1, strFolder2, FileName, CurrentTime, LogMessage, LogFile
    CurrentTime = Year(Now)&"-"&Month(Now)&"-"&Day(Now)&"-"& Hour(Now)&"-"&Minute(Now)&"-"&Second(Now)
    LogMessage  = InputBox("Please Input Your Log Text") 
    DIR = GetCurrentFolder()  
    
    CreateFolder DIR&"\ZipFolderTmp"
    CreateFolder DIR&"\ZipFolderTmp\"&LAB&CurrentTime
    
    set LogFile = fso.CreateTextFile(DIR&"\ZipFolderTmp\"&LAB&CurrentTime&"\BackupLog-"&CurrentTime&".txt", true)
    LogFile.WriteLine(CurrentTime)
    LogFile.WriteBlankLines(1)
    LogFile.WriteLine(LogMessage)
    LogFile.Close()
    
    CreateFolder DIR&"\ZipFolderTmp\"&LAB&CurrentTime&"\mycpu_verify"
    CreateFolder DIR&"\ZipFolderTmp\"&LAB&CurrentTime&"\mycpu_verify\rtl"
    CopyFolder   DIR&"\mycpu_verify\rtl", DIR&"\ZipFolderTmp\"&LAB&CurrentTime&"\mycpu_verify\rtl"
    CreateFolder DIR&"\ZipFolderTmp\"&LAB&CurrentTime&"\mycpu_verify\testbench"
    CopyFolder   DIR&"\mycpu_verify\testbench", DIR&"\ZipFolderTmp\"&LAB&CurrentTime&"\mycpu_verify\testbench"
    
    If fso.FolderExists(DIR & "\PRJ_BACKUP")=False Then         
        CreateFolder DIR&"\PRJ_BACKUP"
    End If
    
    CreateZip    DIR&"\ZipFolderTmp", DIR&"\"&LAB&CurrentTime&".zip"
    wScript.Sleep 5000
    CopyFile     DIR&"\"&LAB&CurrentTime&".zip", DIR&"\PRJ_BACKUP"
    DeleteFolder DIR & "\ZipFolderTmp"
    DeleteFile   DIR&"\"&LAB&CurrentTime&".zip"
    WScript.Quit
End Function

Function GetFileName(srcFile)
    GetFileName = right(srcFile, len(srcFile)-(instrrev(srcFile,"\")))
End Function

Function FileRename(srcFile, NewName)
    set File = fso.getfile(srcFile)
    File.name = NewName
End Function

Function CreateFolder(dstFolder)
    fso.CreateFolder dstFolder
End Function

Function DeleteFolder(dstFolder)
    fso.DeleteFolder dstFolder, True
End Function

Function DeleteFile(FileName)
    fso.DeleteFile Filename, True
End Function

Function CopyFolder(srcFolder, dstFolder)
    Dim Folder,subFolders,Files,File
    Set Folder = fso.Getfolder(srcFolder)
    Set subFolders = Folder.subFolders
    Set Files = Folder.Files
    For Each File In Files
        fso.CopyFile File.Path, dstFolder & "\", False 
        If Err.Number<>0 Then Err.Clear
    Next
    For Each subfolder In subFolders
        CreateFolder(dstFolder & "\" & subFolder.Name)
        call CopyFolder(subFolder.Path, dstFolder & "\" & subFolder.Name)
    Next
End Function

Function CopyFile(srcFile, dstFolder)
    fso.CopyFile srcFile, dstFolder & "\", False 
    If Err.Number<>0 Then Err.Clear
End Function

Function GetCurrentFolder()
    Dim CurrentFile
    CurrentFile = wscript.scriptfullname
    GetCurrentFolder = left(wscript.scriptfullname,instrrev(CurrentFile,"\")-1)
End Function

Function SelectFile()
    Dim FilePath
    Set wShell=CreateObject("WScript.Shell")
    Set oExec=wShell.Exec("mshta.exe ""about:<input type=file id=FILE><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")
    FilePath = CStr(oExec.StdOut.ReadAll)
    SelectFile = Left(FilePath,Len(FilePath)-2)
End Function

Function CreateZip(srcFolder, ZipName)
    Set FS = CreateObject("Scripting.FileSystemObject")
    InputFolder = FS.GetAbsolutePathName(srcFolder)
    ZipFile = FS.GetAbsolutePathName(ZipName)
    CreateObject("Scripting.FileSystemObject").CreateTextFile(ZipFile, True).Write "PK" & Chr(5) & Chr(6) & String(18, vbNullChar)
    Set objShell = CreateObject("Shell.Application")
    Set source = objShell.NameSpace(InputFolder).Items
    objShell.NameSpace(ZipFile).CopyHere(source)
    wScript.Sleep 4000
End Function
