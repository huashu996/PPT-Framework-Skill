param(
    [Parameter(Mandatory=$true)][string]$MathMlPath,
    [Parameter(Mandatory=$true)][int]$MathTypeProcessId,
    [ValidateRange(1,30)][int]$TimeoutSeconds=8
)
$ErrorActionPreference='Stop'
if([Threading.Thread]::CurrentThread.ApartmentState-ne'STA'){throw 'MathType editor automation requires the single STA pipeline host.'}
$path=[IO.Path]::GetFullPath($MathMlPath);if(-not(Test-Path -LiteralPath $path)){throw "MathML file does not exist: $path"}
$expectedMathMl=[IO.File]::ReadAllText($path,[Text.Encoding]::UTF8)
$process=Get-Process -Id $MathTypeProcessId -ErrorAction Stop
if($process.MainWindowHandle-eq 0){throw "MathType process $MathTypeProcessId has no visible editor window."}
if(-not('FormulaSkillEditorWindow'-as[type])){Add-Type @'
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class FormulaSkillEditorWindow {
 [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
 [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd,int nCmdShow);
 [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr hWnd);
 [DllImport("user32.dll")] public static extern IntPtr SetFocus(IntPtr hWnd);
 [DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
 [DllImport("user32.dll")] public static extern bool AttachThreadInput(uint idAttach,uint idAttachTo,bool attach);
 [DllImport("user32.dll")] public static extern void SwitchToThisWindow(IntPtr hWnd,bool altTab);
 [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd,out uint pid);
 [DllImport("user32.dll")] public static extern IntPtr GetKeyboardLayout(uint idThread);
 [DllImport("user32.dll")] public static extern IntPtr LoadKeyboardLayout(string pwszKLID,uint Flags);
 [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr hWnd,uint Msg,IntPtr wParam,IntPtr lParam);
 [DllImport("user32.dll")] public static extern uint GetClipboardSequenceNumber();
 [DllImport("user32.dll")] public static extern IntPtr GetLastActivePopup(IntPtr hWnd);
 [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr hWnd,StringBuilder text,int maxCount);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr hWnd,StringBuilder text,int maxCount);
 public delegate bool EnumWindowsProc(IntPtr hWnd,IntPtr lParam);
 [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr hWnd,EnumWindowsProc callback,IntPtr lParam);
 [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
 [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd,out RECT rect);
 [DllImport("user32.dll")] public static extern bool SetCursorPos(int x,int y);
 [DllImport("user32.dll")] public static extern void mouse_event(uint flags,uint dx,uint dy,uint data,UIntPtr extraInfo);
 [StructLayout(LayoutKind.Sequential)] public struct INPUT { public uint type; public InputUnion U; }
 [StructLayout(LayoutKind.Explicit)] public struct InputUnion {
   [FieldOffset(0)] public MOUSEINPUT mi;
   [FieldOffset(0)] public KEYBDINPUT ki;
   [FieldOffset(0)] public HARDWAREINPUT hi;
 }
 [StructLayout(LayoutKind.Sequential)] public struct MOUSEINPUT { public int dx; public int dy; public uint mouseData; public uint dwFlags; public uint time; public UIntPtr dwExtraInfo; }
 [StructLayout(LayoutKind.Sequential)] public struct KEYBDINPUT { public ushort wVk; public ushort wScan; public uint dwFlags; public uint time; public UIntPtr dwExtraInfo; }
 [StructLayout(LayoutKind.Sequential)] public struct HARDWAREINPUT { public uint uMsg; public ushort wParamL; public ushort wParamH; }
 [DllImport("user32.dll",SetLastError=true)] public static extern uint SendInput(uint count,INPUT[] inputs,int size);
 public static uint SendKeyEvent(ushort key,bool keyUp) {
   INPUT[] inputs=new INPUT[1];
   inputs[0].type=1;inputs[0].U.ki.wVk=key;inputs[0].U.ki.dwFlags=keyUp?2u:0u;
   return SendInput(1,inputs,Marshal.SizeOf(typeof(INPUT)));
 }
 public static bool ClickEditorPane(IntPtr hWnd) {
   RECT rect;
   if(!GetWindowRect(hWnd,out rect)) return false;
   int width=Math.Max(1,rect.Right-rect.Left);
   int height=Math.Max(1,rect.Bottom-rect.Top);
   int x=rect.Left+Math.Min(300,Math.Max(100,width/5));
   int y=rect.Top+Math.Min(height-80,Math.Max(460,(int)(height*0.58)));
   if(!SetCursorPos(x,y)) return false;
   mouse_event(0x0002,0,0,0,UIntPtr.Zero);
   mouse_event(0x0004,0,0,0,UIntPtr.Zero);
   return true;
 }
 public static bool AcknowledgeModal(IntPtr owner) {
   IntPtr popup=GetLastActivePopup(owner);
   if(popup==IntPtr.Zero || popup==owner || !IsWindowVisible(popup)) return false;
   StringBuilder cls=new StringBuilder(64);
   GetClassName(popup,cls,cls.Capacity);
   if(cls.ToString()!="#32770") return false;
   SetForegroundWindow(popup);
   BringWindowToTop(popup);
   SendKeyEvent(0x0D,false);
   SendKeyEvent(0x0D,true);
   return true;
   /*
   IntPtr okButton=IntPtr.Zero;
   EnumChildWindows(popup,delegate(IntPtr child,IntPtr unused) {
     StringBuilder childClass=new StringBuilder(64);
     GetClassName(child,childClass,childClass.Capacity);
     if(childClass.ToString()!="Button") return true;
     StringBuilder caption=new StringBuilder(128);
     GetWindowText(child,caption,caption.Capacity);
     string label=caption.ToString().Replace("&","").Trim();
     if(label=="确定" || label.Equals("OK",StringComparison.OrdinalIgnoreCase)) {
       okButton=child;
       return false;
     }
     return true;
   },IntPtr.Zero);
   if(okButton!=IntPtr.Zero) {
     SendMessage(okButton,0x00F5,IntPtr.Zero,IntPtr.Zero);
     return true;
   }
   SendMessage(popup,0x0111,new IntPtr(1),IntPtr.Zero);
   return true;
   */
 }
}
'@}
function Send-ControlKey([uint16]$KeyCode){
    if([FormulaSkillEditorWindow]::SendKeyEvent(0x11,$false)-ne1){throw 'Failed to press Ctrl.'}
    try{
        Start-Sleep -Milliseconds 35
        if([FormulaSkillEditorWindow]::SendKeyEvent($KeyCode,$false)-ne1){throw "Failed to press key 0x$($KeyCode.ToString('X2'))."}
        Start-Sleep -Milliseconds 25
        if([FormulaSkillEditorWindow]::SendKeyEvent($KeyCode,$true)-ne1){throw "Failed to release key 0x$($KeyCode.ToString('X2'))."}
    }
    finally{
        Start-Sleep -Milliseconds 25
        [void][FormulaSkillEditorWindow]::SendKeyEvent(0x11,$true)
    }
    Start-Sleep -Milliseconds 35
}
$shell=New-Object -ComObject WScript.Shell
$handle=$process.MainWindowHandle
function Close-MathTypePasteErrorDialogIfPresent {
    if([FormulaSkillEditorWindow]::AcknowledgeModal($handle)){
        Start-Sleep -Milliseconds 220
        return $true
    }
    $localizedTitle='MathType '+[string]([char[]](0x9519,0x8BEF))
    foreach($title in @($localizedTitle,'MathType Error')){
        if($shell.AppActivate($title)){
            Start-Sleep -Milliseconds 120
            $shell.SendKeys('{ENTER}')
            Start-Sleep -Milliseconds 220
            return $true
        }
    }
    return $false
}
function Focus-MathTypeEditor {
    $foregroundHandle=[FormulaSkillEditorWindow]::GetForegroundWindow();$activeProcessId=[uint32]0
    $foregroundThread=[FormulaSkillEditorWindow]::GetWindowThreadProcessId($foregroundHandle,[ref]$activeProcessId)
    $currentThread=[FormulaSkillEditorWindow]::GetCurrentThreadId()
    [void][FormulaSkillEditorWindow]::AttachThreadInput($currentThread,$foregroundThread,$true)
    try {
        [void][FormulaSkillEditorWindow]::ShowWindow($handle,9)
        [void][FormulaSkillEditorWindow]::BringWindowToTop($handle)
        [void][FormulaSkillEditorWindow]::SetForegroundWindow($handle)
        [void][FormulaSkillEditorWindow]::SetFocus($handle)
        [FormulaSkillEditorWindow]::SwitchToThisWindow($handle,$true)
        [void]$shell.AppActivate($MathTypeProcessId)
    }
    finally { [void][FormulaSkillEditorWindow]::AttachThreadInput($currentThread,$foregroundThread,$false) }
    $focusDeadline=[DateTime]::UtcNow.AddSeconds($TimeoutSeconds);$foregroundProcessId=[uint32]0
    do{
        Start-Sleep -Milliseconds 60
        $foregroundWindow=[FormulaSkillEditorWindow]::GetForegroundWindow()
        [void][FormulaSkillEditorWindow]::GetWindowThreadProcessId($foregroundWindow,[ref]$foregroundProcessId)
    }while($foregroundProcessId-ne$MathTypeProcessId -and [DateTime]::UtcNow-lt$focusDeadline)
    if($foregroundProcessId-ne$MathTypeProcessId){throw "MathType lost foreground focus before keyboard input: expected $MathTypeProcessId, got $foregroundProcessId."}
    if(-not[FormulaSkillEditorWindow]::ClickEditorPane($handle)){throw 'Failed to focus the MathType equation editing pane.'}
    Start-Sleep -Milliseconds 120
}
Focus-MathTypeEditor
if(Close-MathTypePasteErrorDialogIfPresent){
    Focus-MathTypeEditor
}
$foregroundPid=[uint32]0
$threadId=[FormulaSkillEditorWindow]::GetWindowThreadProcessId($handle,[ref]$foregroundPid)
$layout=[FormulaSkillEditorWindow]::GetKeyboardLayout($threadId);$languageId=([int64]$layout)-band 0xFFFF
if($languageId-ne0x0409){$english=[FormulaSkillEditorWindow]::LoadKeyboardLayout('00000409',1);[void][FormulaSkillEditorWindow]::SendMessage($handle,0x0050,[IntPtr]::Zero,$english);Start-Sleep -Milliseconds 120;$layout=[FormulaSkillEditorWindow]::GetKeyboardLayout($threadId);$languageId=([int64]$layout)-band 0xFFFF}
if($languageId-ne0x0409){throw 'MathType input method is not English/ASCII.'}

# Complete editor preparation before replacing the clipboard. Once the current
# MathML is on the clipboard, do not switch windows or click again before Ctrl+V.
Focus-MathTypeEditor
Send-ControlKey 0x41
Start-Sleep -Milliseconds 100
& (Join-Path $PSScriptRoot 'set_mathtype_mathml_clipboard.ps1') -MathMlPath $path|Out-Null
Add-Type -AssemblyName System.Windows.Forms
$importClipboard=[Windows.Forms.Clipboard]::GetDataObject()
$importFormats=@($importClipboard.GetFormats($false))
if($importFormats -notcontains 'MathML Presentation'){
    throw 'MathType import clipboard is missing MathML Presentation; Ctrl+V was not sent.'
}
$importPayload=$importClipboard.GetData('MathML Presentation',$false)
if($importPayload-is[IO.MemoryStream]){
    $importPayload.Position=0
    $bytes=New-Object byte[] ([int]$importPayload.Length)
    [void]$importPayload.Read($bytes,0,$bytes.Length)
    $importMathMl=[Text.Encoding]::UTF8.GetString($bytes).TrimEnd([char]0)
}
elseif($importPayload-is[byte[]]){
    $importMathMl=[Text.Encoding]::UTF8.GetString($importPayload).TrimEnd([char]0)
}
else{
    $importMathMl=[string]$importPayload
}
if([string]::IsNullOrWhiteSpace($importMathMl)-or$importMathMl-notmatch'<math(?:\s|>)'){
    throw 'MathType import clipboard has no non-empty Presentation MathML; Ctrl+V was not sent.'
}
$importClipboardSequence=[FormulaSkillEditorWindow]::GetClipboardSequenceNumber()
Send-ControlKey 0x56
$importWaitMs=[Math]::Min(3000,[Math]::Max(900,$expectedMathMl.Length*2))
$importDeadline=[DateTime]::UtcNow.AddMilliseconds($importWaitMs)
do{
    Start-Sleep -Milliseconds 100
    if(Close-MathTypePasteErrorDialogIfPresent){
        throw 'MathType rejected the imported clipboard equation; the error dialog was acknowledged immediately and this formula remains pending.'
    }
}while([DateTime]::UtcNow-lt$importDeadline)
if(Close-MathTypePasteErrorDialogIfPresent){
    throw 'MathType rejected the imported clipboard equation; the error dialog was acknowledged immediately and this formula remains pending.'
}
Focus-MathTypeEditor
Send-ControlKey 0x41;Start-Sleep -Milliseconds 100

# Remove the imported MathML payload before Ctrl+C so stale input data can never
# be mistaken for the equation copied from MathType. Then wait for a new
# clipboard sequence number and for MathType's OLE formats to appear.
$cleared=$false
for($clearAttempt=1;$clearAttempt-le20 -and -not$cleared;$clearAttempt++){
    try{
        [Windows.Forms.Clipboard]::Clear()
        $cleared=$true
    }catch{
        Start-Sleep -Milliseconds 40
    }
}
if(-not$cleared){throw 'Could not clear the imported MathML clipboard before copying the MathType equation.'}
$copyStartSequence=[FormulaSkillEditorWindow]::GetClipboardSequenceNumber()
Send-ControlKey 0x43
$copyDeadline=[DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
$clipboardObject=$null
$formats=@()
$oleReady=$false
do{
    Start-Sleep -Milliseconds 100
    try{
        $currentSequence=[FormulaSkillEditorWindow]::GetClipboardSequenceNumber()
        $clipboardObject=[Windows.Forms.Clipboard]::GetDataObject()
        $formats=@($clipboardObject.GetFormats())
        $oleReady=($currentSequence-ne$copyStartSequence) -and (($formats -contains 'MathType EF') -or ($formats -contains 'Embedded Object'))
    }catch{
        $oleReady=$false
    }
}while(-not$oleReady -and [DateTime]::UtcNow-lt$copyDeadline)
if(-not$oleReady){throw "MathType did not place a new MathType OLE object on the clipboard after sequence $copyStartSequence; formats: $($formats -join ', ')"}

$copiedMathMl=$null
foreach($formatName in @('MathML','MathML Presentation','application/mathml+xml')){
    if($formats -contains $formatName){
        try{
            $copiedPayload=$clipboardObject.GetData($formatName)
            if($copiedPayload-is[IO.MemoryStream]){
                $copiedPayload.Position=0
                $copiedBytes=New-Object byte[] ([int]$copiedPayload.Length)
                [void]$copiedPayload.Read($copiedBytes,0,$copiedBytes.Length)
                $copiedMathMl=[Text.Encoding]::UTF8.GetString($copiedBytes).TrimEnd([char]0)
            }
            elseif($copiedPayload-is[byte[]]){
                $copiedMathMl=[Text.Encoding]::UTF8.GetString($copiedPayload).TrimEnd([char]0)
            }
            else{
                $copiedMathMl=[string]$copiedPayload
            }
        }catch{}
        if(-not[string]::IsNullOrWhiteSpace($copiedMathMl)){break}
    }
}
if([string]::IsNullOrWhiteSpace($copiedMathMl)){
    throw 'MathType clipboard contains an OLE format but no non-empty MathML content; empty formula will not be pasted.'
}
function Get-SemanticMathText([string]$MathMl){
    $plain=[Text.RegularExpressions.Regex]::Replace($MathMl,'<[^>]+>','')
    $plain=[Net.WebUtility]::HtmlDecode($plain)
    return [Text.RegularExpressions.Regex]::Replace($plain,'\s+','')
}
$expectedSemantic=Get-SemanticMathText $expectedMathMl
$copiedSemantic=Get-SemanticMathText $copiedMathMl
if($copiedSemantic.Length-eq0){
    throw 'MathType clipboard contains zero semantic characters; empty formula will not be pasted.'
}
[pscustomobject]@{mathml_file=$path;mathtype_process_id=$MathTypeProcessId;input_language_id=('0x{0:X4}'-f$languageId);mathml_imported=$true;editor_content_ready=$true;ole_clipboard_ready=$true;import_clipboard_sequence=$importClipboardSequence;copy_clipboard_sequence=$currentSequence;expected_semantic_length=$expectedSemantic.Length;copied_semantic_length=$copiedSemantic.Length;clipboard_formats=$formats}|ConvertTo-Json -Depth 5
