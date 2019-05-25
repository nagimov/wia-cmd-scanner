Module Module1

    Sub printUsage()
        Const version = "0.2.1"
        Console.WriteLine("wia-cmd-scanner (version " & version & ")                                      ")
        Console.WriteLine("License GPLv3+: GNU GPL version 3 or later <http://gnu.org/licenses/gpl.html>  ")
        Console.WriteLine("                                                                               ")
        Console.WriteLine("Command-line scanner utility for WIA-compatible scanners                       ")
        Console.WriteLine("Online help, docs & bug reports: <https://github.com/nagimov/wia-cmd-scanner/> ")
        Console.WriteLine("                                                                               ")
        Console.WriteLine("Usage: wia-cmd-scanner [OPTIONS]...                                            ")
        Console.WriteLine("                                                                               ")
        Console.WriteLine("All arguments are mandatory. Arguments must be ordered as follows:             ")
        Console.WriteLine("                                                                               ")
        Console.WriteLine("  /w WIDTH                        width of the scan area, mm                   ")
        Console.WriteLine("  /h HEIGHT                       height of the scan area, mm                  ")
        Console.WriteLine("  /dpi RESOLUTION                 scan resolution, dots per inch               ")
        Console.WriteLine("  /color {RGB,GRAY,BW}            scan color mode                              ")
        Console.WriteLine("  /depth {1,8,24}                 scan color depth                              ")
        Console.WriteLine("  /format {BMP,PNG,GIF,JPG,TIF}   output image format                          ")
        Console.WriteLine("  /output FILEPATH                path to output image file                    ")
        Console.WriteLine("                                                                               ")
        Console.WriteLine("Use /w 0 and /h 0 for scanner-defined values, e.g. for receipt scanners        ")
        Console.WriteLine("                                                                               ")
        Console.WriteLine("e.g. for A4 size black and white scan at 300 dpi:                              ")
        Console.WriteLine("wia-cmd-scanner /w 210 /h 297 /dpi 300 /color BW /depth 1 /format PNG /output .\scan.png")
    End Sub

    Sub printExceptionMessage(ex As Exception)
        Dim exceptionDesc As New Dictionary(Of String, String)()
        ' see https://docs.microsoft.com/en-us/windows/desktop/wia/-wia-error-codes
        exceptionDesc.Add("0x80210006", "Scanner is busy")
        exceptionDesc.Add("0x80210016", "Cover is open")
        exceptionDesc.Add("0x8021000A", "Can't communicate with scanner")
        exceptionDesc.Add("0x8021000D", "Scanner is locked")
        exceptionDesc.Add("0x8021000E", "Exception in driver occured")
        exceptionDesc.Add("0x80210001", "Unknown error occured")
        exceptionDesc.Add("0x8021000C", "Incorrect scanner setting")
        exceptionDesc.Add("0x8021000F", "Unsupported scanner command")
        exceptionDesc.Add("0x80210009", "Scanner is deleted and no longer available")
        exceptionDesc.Add("0x80210017", "Scanner lamp is off")
        exceptionDesc.Add("0x80210021", "Maximum endorser value reached")
        exceptionDesc.Add("0x80210020", "Multiple page feed error")
        exceptionDesc.Add("0x80210005", "Scanner is offline")
        exceptionDesc.Add("0x80210003", "No document in document feeder")
        exceptionDesc.Add("0x80210002", "Paper jam in document feeder")
        exceptionDesc.Add("0x80210004", "Unspecified error with document feeder")
        exceptionDesc.Add("0x80210007", "Scanner is warming up")
        exceptionDesc.Add("0x80210008", "Unknown problem with scanner")
        exceptionDesc.Add("0x80210015", "No scanners found")
        For Each pair As KeyValuePair(Of String, String) In exceptionDesc
            If ex.Message.Contains(pair.Key) Then
                Console.WriteLine(pair.Value)
                Exit Sub
            End If
        Next
        Console.WriteLine("Exception occured: " & ex.Message)
    End Sub


    Sub Main()
        ' parse command line arguments
        Dim clArgs() As String = Environment.GetCommandLineArgs()

        If (clArgs.Length < 8) Then
            printUsage()
            Exit Sub
        End If

        If Not (clArgs(1) = "/w" And clArgs(3) = "/h" And clArgs(5) = "/dpi" And clArgs(7) = "/color" And clArgs(9) = "/depth" And clArgs(11) = "/format" And clArgs(13) = "/output") Then
            printUsage()
            Exit Sub
        End If

        ' receive cmd line parameters
        Dim w As Double = clArgs(2)
        Dim h As Double = clArgs(4)
        Dim dpi As Integer = clArgs(6)
        Dim color As String = clArgs(8)
        Dim depth As Integer = clArgs(10)
        Dim format As String = clArgs(12)
        Dim output As String = clArgs(14)

        If Not ((w = 0 And h = 0) Or (w > 0 And h > 0)) Then
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
        Try

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
                        ' set scan parameters
                        Dim props As New Dictionary(Of String(), Double)()
                        props.Add({"4104", "WIA_IPA_DEPTH", "color depth"}, depth) ' color mode
                        props.Add({"6146", "WIA_IPS_CUR_INTENT", "color mode"}, colorcode) ' color mode
                        props.Add({"6147", "WIA_IPS_XRES", "resolution"}, dpi) ' horizontal dpi
                        props.Add({"6148", "WIA_IPS_YRES", "resolution"}, dpi) ' vertical dpi
                        If w > 0 Then
                            props.Add({"6151", "WIA_IPS_XEXTENT", "width"}, w / 25.4 * dpi) ' width in pixels
                            props.Add({"6152", "WIA_IPS_YEXTENT", "height"}, h / 25.4 * dpi) ' height in pixels
                            props.Add({"6149", "WIA_IPS_XPOS", "exit"}, 0) ' x origin of scan area
                            props.Add({"6150", "WIA_IPS_YPOS", "exit"}, 0) ' y origin of scan area
                        End If
                        For Each pair As KeyValuePair(Of String(), Double) In props
                            Try
                                With Scanner.Items(1)
                                    .Properties(pair.Key(0)).Value = pair.Value
                                End With
                            Catch ex As Exception
                                Console.WriteLine("Can't set property " & pair.Key(1))
                                If (pair.Key(2) <> "exit") Then
                                    Console.WriteLine("Unsupported parameter, try scanning with different " & pair.Key(2))
                                Else
                                    Console.WriteLine("Unknown issue, quitting")
                                End If
                                Exit Sub
                            End Try
                        Next
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
                    End If
                    Dim TimeEnd = DateAndTime.Second(Now) + (DateAndTime.Minute(Now) * 60) + (DateAndTime.Hour(Now) * 3600)
                    Console.WriteLine("Scan finished in " & (TimeEnd - TimeStart) & " seconds")
                    Exit Sub ' if successfully found and scanned, quit
                End If
            Next
        Catch ex As Exception
            printExceptionMessage(ex)
            Exit Sub
        End Try
    End Sub

End Module
