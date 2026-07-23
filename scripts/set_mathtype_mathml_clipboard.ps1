param(
    [Parameter(Mandatory = $true)][string]$MathMlPath
)

$ErrorActionPreference = 'Stop'
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    throw 'Clipboard access requires STA. Run with: powershell -Sta -File set_mathtype_mathml_clipboard.ps1 ...'
}

$path = [IO.Path]::GetFullPath($MathMlPath)
if (-not (Test-Path -LiteralPath $path)) { throw "MathML file does not exist: $path" }

$mathml = [IO.File]::ReadAllText($path, [Text.Encoding]::UTF8)
if ([string]::IsNullOrWhiteSpace($mathml) -or $mathml -notmatch '<math(?:\s|>)') {
    throw "File does not contain Presentation MathML: $path"
}
try {
    [xml]$mathDocument = $mathml
}
catch {
    throw "File contains malformed MathML XML: $path"
}
$emptyFormulaTokens = @($mathDocument.SelectNodes("//*[local-name()='mi' or local-name()='mo' or local-name()='mn' or local-name()='mtext']") | Where-Object {
    [string]::IsNullOrWhiteSpace($_.InnerText)
})
if ($emptyFormulaTokens.Count -gt 0) {
    throw "MathML contains $($emptyFormulaTokens.Count) empty formula token(s); clipboard was not modified: $path"
}
$semanticText = [Text.RegularExpressions.Regex]::Replace($mathml, '<[^>]+>', '')
$semanticText = [Net.WebUtility]::HtmlDecode($semanticText)
$semanticText = [Text.RegularExpressions.Regex]::Replace($semanticText, '\s+', '')
if ([string]::IsNullOrWhiteSpace($semanticText)) {
    throw "MathML has no visible semantic formula content; clipboard was not modified: $path"
}

if (-not ('FormulaSkillNativeClipboard' -as [type])) {
    Add-Type @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

public static class FormulaSkillNativeClipboard {
    const uint GMEM_MOVEABLE = 0x0002;
    const uint CF_UNICODETEXT = 13;

    [DllImport("user32.dll", SetLastError=true)] static extern bool OpenClipboard(IntPtr hWndNewOwner);
    [DllImport("user32.dll", SetLastError=true)] static extern bool CloseClipboard();
    [DllImport("user32.dll", SetLastError=true)] static extern bool EmptyClipboard();
    [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
    static extern uint RegisterClipboardFormat(string lpszFormat);
    [DllImport("user32.dll", SetLastError=true)] static extern IntPtr SetClipboardData(uint uFormat, IntPtr hMem);
    [DllImport("kernel32.dll", SetLastError=true)] static extern IntPtr GlobalAlloc(uint uFlags, UIntPtr dwBytes);
    [DllImport("kernel32.dll", SetLastError=true)] static extern IntPtr GlobalLock(IntPtr hMem);
    [DllImport("kernel32.dll", SetLastError=true)] static extern bool GlobalUnlock(IntPtr hMem);
    [DllImport("kernel32.dll", SetLastError=true)] static extern IntPtr GlobalFree(IntPtr hMem);

    static IntPtr Allocate(byte[] bytes) {
        IntPtr hMem = GlobalAlloc(GMEM_MOVEABLE, (UIntPtr)bytes.Length);
        if (hMem == IntPtr.Zero) throw new Win32Exception(Marshal.GetLastWin32Error(), "GlobalAlloc failed.");
        IntPtr ptr = GlobalLock(hMem);
        if (ptr == IntPtr.Zero) {
            GlobalFree(hMem);
            throw new Win32Exception(Marshal.GetLastWin32Error(), "GlobalLock failed.");
        }
        Marshal.Copy(bytes, 0, ptr, bytes.Length);
        GlobalUnlock(hMem);
        return hMem;
    }

    static void Put(uint format, byte[] bytes) {
        IntPtr hMem = Allocate(bytes);
        if (SetClipboardData(format, hMem) == IntPtr.Zero) {
            GlobalFree(hMem);
            throw new Win32Exception(Marshal.GetLastWin32Error(), "SetClipboardData failed.");
        }
        // Clipboard owns hMem after a successful SetClipboardData call.
    }

    public static void SetMathMl(string mathml) {
        bool opened = false;
        for (int attempt = 0; attempt < 20 && !opened; attempt++) {
            opened = OpenClipboard(IntPtr.Zero);
            if (!opened) Thread.Sleep(25);
        }
        if (!opened) throw new Win32Exception(Marshal.GetLastWin32Error(), "OpenClipboard failed.");
        try {
            if (!EmptyClipboard()) throw new Win32Exception(Marshal.GetLastWin32Error(), "EmptyClipboard failed.");

            // MathType's documented clipboard representation is UTF-8 and NULL terminated.
            byte[] utf8 = Encoding.UTF8.GetBytes(mathml + "\0");
            foreach (string name in new[] { "MathML Presentation", "MathML", "application/mathml+xml" }) {
                uint format = RegisterClipboardFormat(name);
                if (format == 0) throw new Win32Exception(Marshal.GetLastWin32Error(), "RegisterClipboardFormat failed: " + name);
                Put(format, utf8);
            }

            byte[] unicode = Encoding.Unicode.GetBytes(mathml + "\0");
            Put(CF_UNICODETEXT, unicode);
        }
        finally {
            CloseClipboard();
        }
    }
}
'@
}

[FormulaSkillNativeClipboard]::SetMathMl($mathml)

[pscustomobject]@{
    path = $path
    characters = $mathml.Length
    encoding = 'UTF-8'
    null_terminated = $true
    formats = @('MathML Presentation', 'MathML', 'application/mathml+xml', 'UnicodeText')
} | ConvertTo-Json -Depth 4
