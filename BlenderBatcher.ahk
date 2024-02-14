/*

====================================================================

		BlenderBatcher.ahk	Version: 0.1

		By: Nidhal Flowgun (BoubakerNidhal@hotmail.com)

====================================================================
	
	Description: This script allows to select one Python Script and a directory that has blend files
	and it will automatically run the selected Python script on all blend files in the selected directory
	and all of its subfolders.
	
	The script allows to resume the progress of any chosen Python script, instead of having to restart the operations
	on all the blend files. It does this by storing the name of the completed files in a TXT file that corresponds to
	the Python script. The TXT is saved where this autohotkey script is.
	
	The Python script is run on all the found blend files in the background. It uses the Blender installation that has its path
	selected. This Autohotkey script provides an option to run Blender without any of its addons to speed up the process.
	It does this by renaming the Addon folder so that Blender doesn't find it, and when the operation is complete or aborted,
	it renames it back to its original format.


====================================================================
*/
ListLines Off
#NoEnv
#Requires AutoHotkey v1
SetBatchLines -1
#SingleInstance Force
#Persistent
SetWorkingDir %A_ScriptDir%
DetectHiddenWindows, On
#MaxMem 4095
Process, Priority,, High

if (!A_IsAdmin) {
	Run % "*RunAs " A_ScriptFullPath
	ExitApp
}

OnExit("Exiting") 

if (FileExist(A_ProgramFiles "\Blender Foundation")){
	loop, Files, %A_ProgramFiles%\Blender Foundation\*, D
		{
		Version:=StrReplace(A_LoopFileName,"Blender ","")
		if (VerCompare(Version, StartVer)>0)
			StartVer:= Version
		}
	BlenderDirectory:= A_ProgramFiles "\Blender Foundation\Blender " Version
}

Gui, MainInterface:New, AlwaysOnTop OwnDialogs +ToolWindow -DPIScale +HwndMainInterfaceHwnd

Gui Add, Text, x16 y24 w120 h23 +0x200, Blender Directory:
Gui Add, Edit, x16 y56 w347 h21	vBlenderDirectory, %BlenderDirectory%
Gui Add, Button, x376 y56 w80 h23 vBlenderFolder, Browse
GuiBinder:= Func("FolderSelect").Bind("Blender")
GuiControl, +g, BlenderFolder , % GuiBinder

Gui Add, Text, x16 y104 w120 h23 +0x200, Python Script Directory:
Gui Add, Edit, x16 y136 w347 h21 vPythonScriptDirectory
Gui Add, Button, x376 y136 w80 h23 vPythonScriptFolder gSelectScript, Browse

Gui Add, Text, x16 y184 w120 h23 +0x200, Blend Files Directory:
Gui Add, Edit, x16 y216 w347 h21 vBlendFilesDirectory
Gui Add, Button, x376 y216 w80 h23 vBlendFilesFolder, Browse
GuiBinder:= Func("FolderSelect").Bind("BlendFiles")
GuiControl, +g, BlendFilesFolder , % GuiBinder

Gui Add, CheckBox, x40 y260 vCleanBlender Checked1,Run Blender in the Background without installed addons

Gui Add, Button, x48 y290 w115 h42 gConfirm, Confirm
Gui Add, Button, x248 y290 w115 h42 gExit, Exit
Gui Add, Button, x344 y352 w114 h28 gAbout, About

Gui Show, w469 h392, Window

return

Exit:
ExitApp
return

Exiting(){
	if (BackupDest){
	if WinExist("ahk_exe Blender.exe"){
		msgbox, 4096,Please Close Blender, Complete!`nWe detected that you have Blender open.`nPlease close it so that we restore the addons.
		WinwaitClose, ahk_exe Blender.exe
		FileMoveDir, %BackupDest%,%AddonFolder%, R
	}
	FileMoveDir, %BackupDest%,%AddonFolder%, R
	}
	return
}

About:
Msgbox, 4096,,Script Written By Nidhal Flowgun.`nBoubakerNidhal@hotmail.com
return

SelectScript:
Gui +OwnDialogs
Gui, Submit , NoHide
if (PythonScriptDirectory and FileExist(PythonScriptDirectory)){
	SplitPath,PythonScriptDirectory,,PythonDir
	FileSelectFile, PythonScript, 3,%PythonDir%, Select Python Script, Python Files (*.py)
}else
	FileSelectFile, PythonScript, 3,, Select Python Script, Python Files (*.py)
if (PythonScript){
	GuiControl,, PythonScriptDirectory, %PythonScript%
	Gui, Submit , NoHide
}
return

FolderSelect(FolderType){
Global
local NewDirectory
Gui, Submit , NoHide
if (%Foldertype%Directory and FileExist(Foldertype "Directory")){
	InitialDir:= %Foldertype%Directory
}else if (FolderType="Blender" and FileExist(A_ProgramFiles "\Blender Foundation"))
	InitialDir:= A_ProgramFiles "\Blender Foundation"
else
	InitialDir:= A_ScriptDir

	NewDirectory := ChooseFolder( [MainInterfaceHwnd, "Select " FolderType " Folder"]
						   , InitialDir
						   , [A_ProgramFiles,A_Desktop,A_Temp,A_Startup]
						   , 0x10000000 | 0x02000000 )

	NewDirectory:= Trim(NewDirectory," `t`r`n")
	if (NewDirectory and FileExist(NewDirectory)) {	;; if we didn't cancel adding a new directory or if we gave an argument
		Gui, MainInterface:Default
		if (FolderType="Blender"){
			BlenderFolder:= NewDirectory         
			GuiControl,, BlenderDirectory, %NewDirectory%
		}else if (FolderType="BlendFiles"){
			BlendFilesFolder:= NewDirectory
			GuiControl,, BlendFilesDirectory, %NewDirectory%
		}
		Gui, Submit , NoHide
	}
return
}

Confirm:
Gui, Submit , NoHide
BlenderDirectory:= Trim(BlenderDirectory," `t`r`n""\")
BlendFilesDirectory:= Trim(BlendFilesDirectory," `t`r`n""\")
PythonScriptDirectory:= Trim(PythonScriptDirectory," `t`r`n""\")

if (!FileExist(PythonScriptDirectory)){
	Msgbox, 4096,py File missing, The python Script directory is not valid.
	return
}else{
	SplitPath,PythonScriptDirectory,,,Extension,ScriptFileName
	if (Extension !="py"){
		Msgbox, 4096,py File missing, The python Script directory is not valid.
		return	
	}
}
if (!FileExist(BlendFilesDirectory)){
	Msgbox, 4096,, The directory for Blend files does not exist.
	return
}
if !FileExist(BlenderDirectory "\blender.exe"){
	Msgbox, 4096,, The Blender Directory is not valid.`nIt should be where blender.exe is
	return
}

if (CleanBlender){
if WinExist("ahk_exe Blender.exe"){
	Msgbox,4096,Close Blender, ""Run Blender in the Background without installed addons"" is checked.`nPlease Close All instances of Blender To Continue.`n`nThis option needs Blender closed so that it renames the addons folder to run Blender in the background without the addons, then it restores it back when finishing or exiting this script.
	return
}
	Version:=Substr(BlenderDirectory,-2)
	AddonFolder:= A_Appdata "\Blender Foundation\Blender\" Version "\scripts\addons"
	if FileExist(AddonFolder){
	BackupDest:= A_Appdata "\Blender Foundation\Blender\" Version "\scripts\addons_BAK"
	While FileExist(BackupDest){
		Suffix++
		BackupDest:= A_Appdata "\Blender Foundation\Blender\" Version "\scripts\addons_BAK_" Suffix
	}
	FileMoveDir, %AddonFolder%, %BackupDest%, R
	}
}


return
trackerFile:= ScriptFileName ".txt"
If FileExist(trackerFile){
Msgbox,  4131,Resume?, Do you want to resume from previous progress?
IfMsgBox, Cancel
	return
IfMsgBox, No
	FileDelete, %trackerFile%
IfMsgBox, Yes
	FileRead, CompletedFiles, %trackerFile%
}

SysGet, MonitorPrimary, MonitorPrimary
SysGet, Mon, MonitorWorkArea, %MonitorPrimary%
CoordMode, Tooltip, Screen
Tooltip, We are running the script.`nThis can take a while`nPlease wait...,MonRight-250,MonBottom-100

Loop, Files, %BlendFilesDirectory%\*.blend, F R
	{
	if (CompletedFiles and (instr(CompletedFiles,A_LoopFileFullPath)))
		continue
	Tooltip, We are running the script.`nThis can take a while`nPlease wait...`n%A_loopFileName%,MonRight-250,MonBottom-100
	RunThis:= blender -b "C:\Blender Files\Asset Browser Library\Interior Models\Interior_models_VOL_01\Beds_VOL_01\bed_03_01.blend" --python "C:\Blender Files\Asset Browser Library\PreviewGenFixer.py"
	RunThis:= "blender -b " """" A_LoopFileFullPath """" " --python " """" PythonScriptDirectory """"
	RunWait , %ComSpec% /c cd /d "C:\Program Files\Blender Foundation\Blender 4.0" && %RunThis% ,, hide UseErrorLevel 
	FileAppend, %A_LoopFileFullPath%`n, %trackerFile%	
	}
FileDelete, %trackerFile%
tooltip,

if (CleanBlender){
	DetectHiddenWindows, on
	if WinExist("ahk_exe Blender.exe"){
		msgbox, 4096,Finishing..., Complete!`nWe detected that you have Blender open.`nPlease close it so that we restore the addons.
		WinwaitClose, ahk_exe Blender.exe
		FileMoveDir, %BackupDest%,%AddonFolder%, R
	}else{
	FileMoveDir, %BackupDest%,%AddonFolder%, R
	}
}else
	msgbox, 4096,, Complete!
ExitApp


; Ã©Ã©Ã©Ã©Ã©Ã©Ã©Ã©Ã©Ã©Ã©Ã©Ã©Ã©Ã©Ã©Ã©Ã©Ã©Ã©Ã©Ã©Ã©Ã
; §Ã§Ã§Ã§Ã§Ã§Ã§Ã§Ã§Ã§Ã§Ã§Ã§Ã§Ã§Ã§Ã§Ã§Ã§Ã§Ã§Ã§Ã§Ã§

ChooseFolder(Owner, StartingFolder := "", CustomPlaces := "", Options := 0){
    ; IFileOpenDialog interface
    ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb775834(v=vs.85).aspx
    local IFileOpenDialog := ComObjCreate("{DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7}", "{D57C7288-D4AD-4768-BE02-9D969532D960}")
        ,           Title := IsObject(Owner) ? Owner[2] . "" : ""
        ,           Flags := 0x20 | Options    ; FILEOPENDIALOGOPTIONS enumeration (https://msdn.microsoft.com/en-us/library/windows/desktop/dn457282(v=vs.85).aspx)
        ,      IShellItem := PIDL := 0         ; PIDL recibe la direcci�n de memoria a la estructura ITEMIDLIST que debe ser liberada con la funci�n CoTaskMemFree
        ,             Obj := {}, foo := bar := ""
    Owner := IsObject(Owner) ? Owner[1] : (WinExist("ahk_id" . Owner) ? Owner : 0)
    CustomPlaces := IsObject(CustomPlaces) || CustomPlaces == "" ? CustomPlaces : [CustomPlaces]


    while (InStr(StartingFolder, "\") && !DirExist(StartingFolder))
        StartingFolder := SubStr(StartingFolder, 1, InStr(StartingFolder, "\",, -1) - 1)
    if ( DirExist(StartingFolder) )
	{
        StrPutVar(StartingFolder, StartingFolderW, "UTF-16")
        DllCall("Shell32.dll\SHParseDisplayName", "UPtr", &StartingFolderW, "Ptr", 0, "UPtrP", PIDL, "UInt", 0, "UInt", 0)
        DllCall("Shell32.dll\SHCreateShellItem", "Ptr", 0, "Ptr", 0, "UPtr", PIDL, "UPtrP", IShellItem)
        ObjRawSet(Obj, IShellItem, PIDL)
        ; IFileDialog::SetFolder method
        ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb761828(v=vs.85).aspx
        DllCall(NumGet(NumGet(IFileOpenDialog+0)+12*A_PtrSize), "Ptr", IFileOpenDialog, "UPtr", IShellItem)
	}

    if ( IsObject(CustomPlaces) )
    {
        local Directory := ""
        For foo, Directory in CustomPlaces    ; foo = index
        {
            foo := IsObject(Directory) ? Directory[2] : 0    ; FDAP enumeration (https://msdn.microsoft.com/en-us/library/windows/desktop/bb762502(v=vs.85).aspx)
            if ( DirExist(Directory := IsObject(Directory) ? Directory[1] : Directory) )
            {
                StrPutVar(Directory, DirectoryW, "UTF-16")
                DllCall("Shell32.dll\SHParseDisplayName", "UPtr", &DirectoryW, "Ptr", 0, "UPtrP", PIDL, "UInt", 0, "UInt", 0)
                DllCall("Shell32.dll\SHCreateShellItem", "Ptr", 0, "Ptr", 0, "UPtr", PIDL, "UPtrP", IShellItem)
                ObjRawSet(Obj, IShellItem, PIDL)
                ; IFileDialog::AddPlace method
                ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb775946(v=vs.85).aspx
                DllCall(NumGet(NumGet(IFileOpenDialog+0)+21*A_PtrSize), "UPtr", IFileOpenDialog, "UPtr", IShellItem, "UInt", foo)
            }
        }
    }

    ; IFileDialog::SetTitle method
    ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb761834(v=vs.85).aspx
    StrPutVar(Title, TitleW, "UTF-16")
    DllCall(NumGet(NumGet(IFileOpenDialog+0)+17*A_PtrSize), "UPtr", IFileOpenDialog, "UPtr", Title == "" ? 0 : &TitleW)

    ; IFileDialog::SetOptions method
    ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb761832(v=vs.85).aspx
    DllCall(NumGet(NumGet(IFileOpenDialog+0)+9*A_PtrSize), "UPtr", IFileOpenDialog, "UInt", Flags)

    ; IModalWindow::Show method
    ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb761688(v=vs.85).aspx
    local Result := []
    if ( !DllCall(NumGet(NumGet(IFileOpenDialog+0)+3*A_PtrSize), "UPtr", IFileOpenDialog, "Ptr", Owner, "UInt") )
    {
        ; IFileOpenDialog::GetResults method
        ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb775831(v=vs.85).aspx
        local IShellItemArray := 0    ; IShellItemArray interface (https://msdn.microsoft.com/en-us/library/windows/desktop/bb761106(v=vs.85).aspx)
        DllCall(NumGet(NumGet(IFileOpenDialog+0)+27*A_PtrSize), "UPtr", IFileOpenDialog, "UPtrP", IShellItemArray)

        ; IShellItemArray::GetCount method
        ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb761098(v=vs.85).aspx
        local Count := 0    ; pdwNumItems
        DllCall(NumGet(NumGet(IShellItemArray+0)+7*A_PtrSize), "UPtr", IShellItemArray, "UIntP", Count)

        local Buffer := ""
        VarSetCapacity(Buffer, 32767 * 2)
        loop % Count
        {
            ; IShellItemArray::GetItemAt method
            ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb761100(v=vs.85).aspx
            DllCall(NumGet(NumGet(IShellItemArray+0)+8*A_PtrSize), "UPtr", IShellItemArray, "UInt", A_Index-1, "UPtrP", IShellItem)

            ; IShellItem::GetDisplayName method
            ; https://docs.microsoft.com/en-us/windows/win32/api/shobjidl_core/nf-shobjidl_core-ishellitem-getdisplayname
            DllCall(NumGet(NumGet(IShellItem+0)+5*A_PtrSize), "Ptr", IShellItem, "Int", 0x80028000, "PtrP", ptr:=0)
            ObjRawSet(Obj, IShellItem, ptr), ObjPush(Result, RTrim(StrGet(ptr,"UTF-16"), "\"))

            if (Result[A_Index] ~= "^::")  ; handle "::{00000000-0000-0000-0000-000000000000}\Documents.library-ms" (library)
            {
            	; SHLoadLibraryFromParsingName
            	; https://docs.microsoft.com/en-us/windows/win32/api/shobjidl_core/nf-shobjidl_core-shloadlibraryfromparsingname
                VarSetCapacity(IID_IShellItem, 16)
                DllCall("Ole32\CLSIDFromString", "Str", "{43826d1e-e718-42ee-bc55-a1e261c37bfe}", "Ptr", &IID_IShellItem)
                DllCall("Shell32\SHCreateItemFromParsingName", "WStr", Result[A_Index], "Ptr", 0, "Ptr", &IID_IShellItem, "PtrP", IShellItem:=0)

                IShellLibrary := ComObjCreate("{d9b3211d-e57f-4426-aaef-30a806add397}", "{11A66EFA-382E-451A-9234-1E0E12EF3085}")
                ; IShellLibrary::LoadLibraryFromItem
                ; https://docs.microsoft.com/en-us/windows/win32/api/shobjidl_core/nn-shobjidl_core-ishelllibrary
                DllCall(NumGet(NumGet(IShellLibrary+0)+3*A_PtrSize), "UPtr", IShellLibrary, "Ptr", IShellItem, "Int", 0)
                IShellLibrary2 := ComObjQuery(IShellLibrary, "{11A66EFA-382E-451A-9234-1E0E12EF3085}")
                ObjRelease(IShellLibrary)
                ObjRelease(IShellItem)

                VarSetCapacity(IID_IShellItemArray, 16)
                DllCall("Ole32\CLSIDFromString", "Str", "{b63ea76d-1f85-456f-a19c-48159efa858b}", "Ptr", &IID_IShellItemArray)
                ; IShellLibrary::GetFolders method
                ; https://docs.microsoft.com/en-us/windows/win32/api/shobjidl_core/nf-shobjidl_core-ishelllibrary-getfolders
                DllCall(NumGet(NumGet(IShellLibrary2+0)+7*A_PtrSize), "UPtr", IShellLibrary2, "Int", 1, "Ptr", &IID_IShellItemArray, "PtrP", IShellItemArray:=0)

                DllCall(NumGet(NumGet(IShellItemArray+0)+8*A_PtrSize), "UPtr", IShellItemArray, "Int", 0, "PtrP", IShellItem0:=0)
                DllCall(NumGet(NumGet(IShellItem0+0)+5*A_PtrSize), "Ptr", IShellItem0, "Int", 0x80028000, "PtrP", ptr:=0)
                Result[A_Index] := StrGet(ptr, "UTF-16")
                DllCall("Ole32\CoTaskMemFree", "Ptr", ptr)
                ObjRelease(IShellItem0)
                ObjRelease(IShellItemArray)
                ObjRelease(IShellLibrary2)
            }
        }

        ObjRelease(IShellItemArray)
    }


    for foo, bar in Obj    ; foo = IShellItem interface (ptr)  |  bar = PIDL structure (ptr)
        ObjRelease(foo), DllCall("Ole32.dll\CoTaskMemFree", "Ptr", bar)
    ObjRelease(IFileOpenDialog)

    return ObjLength(Result) ? ( Options & 0x200 ? Result : Result[1] ) : FALSE
}

DirExist(DirName)
{
    loop Files, % DirName, D
        return A_LoopFileAttrib
}

StrPutVar(string, ByRef var, encoding){
    ; Ensure capacity.
    VarSetCapacity( var, StrPut(string, encoding)
        ; StrPut returns char count, but VarSetCapacity needs bytes.
        * ((encoding="utf-16"||encoding="cp1200") ? 2 : 1) )
    ; Copy or convert the string.
    return StrPut(string, &var, encoding)
} ; https://www.autohotkey.com/docs/commands/StrPut.htm#Examples
;;-----------------------------------------------------------------------
