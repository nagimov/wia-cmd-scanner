Module Module1

    Sub printUsage()
        Const version = "0.1"
        Console.WriteLine("wia-cmd-scanner (version " & version & ")                                     ")
        Console.WriteLine("License GPLv3+: GNU GPL version 3 or later <http://gnu.org/licenses/gpl.html> ")
        Console.WriteLine("                                                                              ")
        Console.WriteLine("Command-line scanner utility for WIA-compatible scanners                      ")
        Console.WriteLine("Online help, docs & bug reports: <https://github.com/nagimov/wia-cmd-scanner/>")
        Console.WriteLine("                                                                              ")
        Console.WriteLine("Usage: wia-cmd-scanner [OPTIONS]...                                           ")
        Console.WriteLine("                                                                              ")
        Console.WriteLine("All arguments are mandatory. Arguments must be ordered as follows:            ")
        Console.WriteLine("                                                                              ")
        Console.WriteLine("  /dpi {150,200,300,600}          scan resolution, dots per inch              ")
        Console.WriteLine("  /color {RGB,GRAY,BW}            scan color mode                             ")
        Console.WriteLine("  /format {BMP,PNG,GIF,JPG,TIF}   output image format                         ")
        Console.WriteLine("  /output FILEPATH                path to output image file                   ")
        Console.WriteLine("                                                                              ")
        Console.WriteLine("e.g.:                                                                         ")
        Console.WriteLine("  wia-cmd-scanner /dpi 300 /color RGB /format PNG /output .\scan.png          ")
    End Sub

    Sub Main()
        ' parse command line arguments
        Dim clArgs() As String = Environment.GetCommandLineArgs()

        If (clArgs.Length < 8) Then
            printUsage()
            Exit Sub
        End If

        If Not (clArgs(1) = "/dpi" And clArgs(3) = "/color" And clArgs(5) = "/format" And clArgs(7) = "/output") Then
            printUsage()
            Exit Sub
        End If

        Dim dpi As Integer = clArgs(2)
        Dim color As String = clArgs(4)
        Dim format As String = clArgs(6)
        Dim output As String = clArgs(8)

        If Not (dpi = 150 Or dpi = 200 Or dpi = 300 Or dpi = 600) Then
            printUsage()
            Exit Sub
        End If

        Dim colorcode As Integer
        If color = "RGB" Then
            colorcode = 1
        ElseIf color = "GRAY" Then
            colorcode = 2
        ElseIf color = "BW" Then
            colorcode = 4
        Else
            printUsage()
            Exit Sub
        End If

        Dim fileformat As String
        Const wiaFormatBMP = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
        Const wiaFormatPNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
        Const wiaFormatGIF = "{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}"
        Const wiaFormatJPG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
        Const wiaFormatTIF = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}"

        If format = "BMP" Then
            fileformat = wiaFormatBMP
        ElseIf format = "PNG" Then
            fileformat = wiaFormatPNG
        ElseIf format = "GIF" Then
            fileformat = wiaFormatGIF
        ElseIf format = "JPG" Then
            fileformat = wiaFormatJPG
        ElseIf format = "TIF" Then
            fileformat = wiaFormatTIF
        Else
            printUsage()
            Exit Sub
        End If

        If System.IO.File.Exists(output) = True Then
            Console.WriteLine("Destination file exists!")
            Exit Sub
        End If

        ' scan the image
        Dim DeviceManager = CreateObject("WIA.DeviceManager")  ' create device manager
        If DeviceManager.DeviceInfos.Count < 1 Then
            Console.WriteLine("No compatible scanners found")
            Exit Sub
        End If
        For i = 1 To DeviceManager.DeviceInfos.Count  ' check all available devices
            If DeviceManager.DeviceInfos(i).Type = 1 Then  ' find first device of type "scanner" (exclude webcams, etc.)
                Dim TimeStart = DateAndTime.Second(Now) + (DateAndTime.Minute(Now) * 60) + (DateAndTime.Hour(Now) * 3600)
                Dim Scanner As WIA.Device = DeviceManager.DeviceInfos(i).connect  ' connect to scanner
                If IsNothing(Scanner) Then
                    Console.WriteLine("Scanner " & i & " not recognized")
                Else
                    Console.WriteLine("Scanning to file " & output & " (dpi = " & dpi & ", color mode '" & color & "', output format '" & format & "')")
                    Try
                        ' set scan parameters
                        With Scanner.Items(1)
                            .Properties("6146").Value = colorcode
                            .Properties("6147").Value = dpi  ' horizontal dpi
                            .Properties("6148").Value = dpi  ' vertical dpi
                        End With
                        ' scan image as BMP...
                        Dim Img As WIA.ImageFile = Scanner.Items(1).Transfer(wiaFormatBMP)
                        ' ...and convert it to desired format
                        Dim ImgProc As Object = CreateObject("WIA.ImageProcess")
                        ImgProc.Filters.Add(ImgProc.FilterInfos("Convert").FilterID)
                        ImgProc.Filters(1).Properties("FormatID") = fileformat
                        ImgProc.Filters(1).Properties("Quality") = 75
                        Img = ImgProc.Apply(Img)
                        ' ...and save it to file
                        Img.SaveFile(output)
                    Catch ex As Exception
                        Console.WriteLine("Exception occured: " & ex.Message)
                        Exit Sub
                    End Try
                End If
                Dim TimeEnd = DateAndTime.Second(Now) + (DateAndTime.Minute(Now) * 60) + (DateAndTime.Hour(Now) * 3600)
                Console.WriteLine("Scan finished in " & (TimeEnd - TimeStart) & " seconds")
                Exit Sub ' if successfully found and scanned, quit
            End If
        Next
    End Sub

End Module
