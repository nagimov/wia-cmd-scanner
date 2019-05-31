# wia-cmd-scanner

This (very) small utility (~30 KB executable) provides an easy to use command-line interface to WIA-compatible scanners for Windows OS. If scanner is accessible using `Windows Fax and Scan` application, it is very likely to be compatible with this tool. Compiled binaries can be downloaded from [Releases](https://github.com/nagimov/wia-cmd-scanner/releases)

The utility is built around WIA (Microsoft Windows Image Acquisition Library v2.0) and requires [Microsoft .NET Framework 4.0](https://www.microsoft.com/en-us/download/details.aspx?id=17718) (it is likely already included in your Windows OS). The utility is portable and requires no installation. Both 32-bit and 64-bit versions of Windows XP ([see the note below](#note-for-windows-xp-users)), Windows Vista, Windows 7, Windows 8/8.1, Windows 10 are supported.

## Usage

```
Usage: wia-cmd-scanner [OPTIONS]...

All arguments are mandatory. Arguments must be ordered as follows:

  /w WIDTH                        width of the scan area, mm
  /h HEIGHT                       height of the scan area, mm
  /dpi RESOLUTION                 scan resolution, dots per inch
  /color {RGB,GRAY,BW}            scan color mode
  /format {BMP,PNG,GIF,JPG,TIF}   output image format
  /output FILEPATH                path to output image file

Use /w 0 and /h 0 for scanner-defined values, e.g. for receipt scanners

e.g. for A4 size black and white scan at 300 dpi:
wia-cmd-scanner /w 210 /h 297 /dpi 300 /color BW /format PNG /output .\scan.png
```

## Build

The utility is compiled using Microsoft Visual Studio 2012 Express.

No Visual Studio project files are provided, since the code can be imported into Visual Studio project from scratch in few easy steps (shown according to Visual Studio 2012 layout):

* `File` -> `New Project`
* `Installed` -> `Templates` -> `Visual Basic` -> `Windows` -> `Console Application`
* copy-paste code from [`wia-cmd-scanner.vb`](https://github.com/nagimov/wia-cmd-scanner/raw/master/wia-cmd-scanner.vb) to `Module1.vb` (empty file will be opened in editor)
* right-click on `ConsoleApplication` in `Solution Explorer` and choose `Add Reference...`
    + choose `COM` -> `Type Libraries`
    + search for `image` and select `Microsoft Windows Image Acquisition Library v2.0`
* compile via `BUILD` -> `Build Solution`

## Scripting and automation

You can build your own automation tools around `wia-cmd-scanner.exe` binary using batch/powershell. E.g. a simple batch job infinitely waiting for key press and scanning to a file with timestamp can be very simply achieved as follows:

```
@setlocal enabledelayedexpansion
:loop
    @echo off
    @for /F "usebackq tokens=1,2 delims==" %%i in (`wmic os get LocalDateTime /VALUE 2^>NUL`) do if '.%%i.'=='.LocalDateTime.' set ldt=%%j
    @set ldt=%ldt:~0,4%-%ldt:~4,2%-%ldt:~6,2%_%ldt:~8,2%-%ldt:~10,2%-%ldt:~12,2%
    @echo on
    wia-cmd-scanner.exe /w 215.9 /h 279.4 /dpi 300 /color RGB /format PNG /output ..\scans\scan_%ldt%.png
    pause
goto loop
```

For more sophisticated automated jobs, check out the source code. The project is very simple and easy to modify to fit your own needs.

## Note for Windows XP users

Since Windows XP only includes legacy WIA v1.0 library, WIA v2.0 needs to be installed in order for this utility to work.

Archive with required files: [wiaautsdk.zip](http://vbnet.mvps.org/files/updates/wiaautsdk.zip)

### WIA v2.0 installation

* Copy the `wiaaut.chm` and `wiaaut.chi` files to your `Help` directory (usually located at `C:\Windows\Help`)
* Copy the `wiaaut.dll` file to your `System32` directory (usually located at `C:\Windows\System32`)
* From a Command Prompt in the `System32` directory run the following command: `RegSvr32 WIAAut.dll`
