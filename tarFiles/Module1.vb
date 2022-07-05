'Imports System.IO
'Imports System.IO.Compression
'Imports System.IO.Compression.fileSystem
Module Module1

    Sub processSubFolder(ByVal srcDir As String, ByVal dstDir As String, ByVal num As Integer, ByVal numDelSrcFiles As Integer, ByRef sumCsvLogWriter As StreamWriter)
        Dim csvLogWriter As System.IO.StreamWriter
        Dim csvLogFile As String = Path.GetFileName(srcDir) + ".csv"

        '若csv存在則忽略這個目錄
        If My.Computer.FileSystem.FileExists(dstDir + "\" + csvLogFile) Then
            Console.WriteLine("{0} csv檔案存在, 略過此目錄", dstDir + "\" + csvLogFile)
            Exit Sub
        End If
        If Directory.Exists(srcDir) Then
            Dim files() As String = Directory.GetFiles(srcDir, "*.zip") '僅檢查.zip的檔案
            Console.WriteLine("{0} 檔案總數: {1}", srcDir, files.Count())

            '計算打包檔案數
            Dim packedNum As Integer = Math.Ceiling(files.Count() / num)
            Dim packedFileName As String
            Dim startPos As Integer
            Dim endPos As Integer
            '最後一個打包檔要特別處理
            Dim endPosLast As Integer = files.Count() - 1
            Console.WriteLine(packedNum.ToString)

            '建立csv log 
            csvLogWriter = My.Computer.FileSystem.OpenTextFileWriter(dstDir + "\" + csvLogFile, True)
            csvLogWriter.WriteLine("原始檔案,目的檔案")


            '每個打包檔流程
            For index As Integer = 1 To packedNum
                '計算檔名
                startPos = (index - 1) * num
                endPos = index * num - 1

                '檢查是否是最後一個打包檔
                If files.Count() - endPos <= 0 Then
                    endPos = endPosLast
                End If

                packedFileName = Path.GetFileName(srcDir) + "_" + index.ToString.PadLeft(12, "0") + ".zip" '當Path是給目錄, GetFileName回傳最後一層目錄名稱
                packedFileName = dstDir + "\" + packedFileName
                Console.WriteLine(packedFileName)
                Console.WriteLine("-----------------------------------------------------------------------------------")

                '產生csv檔案清單
                For i2 As Integer = startPos To endPos
                    'Console.WriteLine(files(i2))
                    csvLogWriter.WriteLine("{0},{1}", files(i2), packedFileName)
                    sumCsvLogWriter.WriteLine("{0},{1}", files(i2), packedFileName)
                Next

                'https://docs.microsoft.com/en-us/dotnet/visual-basic/programming-guide/language-features/arrays/
                Dim filesTmp(endPos - startPos) As String
                Array.Copy(files, startPos, filesTmp, 0, endPos - startPos + 1)
                'For Each aa As String In filesTmp
                '    Console.WriteLine(aa)
                'Next
                tarFile(packedFileName, filesTmp, srcDir)

            Next

            csvLogWriter.Close()

            '檢查是否需刪除來源檔案
            If numDelSrcFiles = 1 Then

                '直接刪目錄 (暫不使用此方式，因files字串陣列僅搜尋*.zip檔案，若直接刪目錄可能會誤刪除未打包檔案)
                'My.Computer.FileSystem.DeleteDirectory(srcDir, FileIO.DeleteDirectoryOption.DeleteAllContents)

                '個別刪除目錄內資料
                'https://docs.microsoft.com/zh-tw/dotnet/visual-basic/developing-apps/programming/drives-directories-files/how-to-delete-a-file
                For Each file As String In files
                    My.Computer.FileSystem.DeleteFile(file, FileIO.UIOption.AllDialogs, FileIO.RecycleOption.SendToRecycleBin)
                    'System.IO.File.Delete(File)
                    Console.WriteLine("Deleting {0}", file)
                Next

            End If

        Else
            Console.WriteLine("來源目錄 [{0}] 不存在!!", srcDir)
            Exit Sub
        End If
    End Sub

    Function tarFile(ByVal zipSavePath As String, ByVal files() As String, ByVal srcDir As String) As Boolean
        'https://docs.microsoft.com/zh-tw/dotnet/standard/io/how-to-compress-and-extract-files
        Using zipToOpen As FileStream = New FileStream(zipSavePath, FileMode.OpenOrCreate)
            Using archive As ZipArchive = New ZipArchive(zipToOpen, ZipArchiveMode.Update)

                For Each file As String In files
                    'https://docs.microsoft.com/en-us/dotnet/framework/migration-guide/mitigation-ziparchiveentry-fullname-path-separator
                    Dim tmpFilePath As String = Path.GetFullPath(file).Substring(srcDir.Length) '移除多餘目錄結構
                    If tmpFilePath.Substring(0, 1) = "\" Then '若開頭有\, 手動移除
                        tmpFilePath = tmpFilePath.Substring(1)
                    End If
                    archive.CreateEntryFromFile(file, tmpFilePath, 2)
                    'archive.CreateEntryFromFile(file, Path.GetFileName(file), 2)
                    'archive.CreateEntryFromFile(file, Path.GetFullPath(file), 2)
                    Console.Write(".")
                Next

            End Using
            Console.WriteLine()

        End Using
        Return True
    End Function
    Sub Main()
        Dim cmd_args() As String = Environment.GetCommandLineArgs()
        If cmd_args.Count() < 6 Then
            ' https://social.msdn.microsoft.com/Forums/vstudio/en-US/51379b07-7fd3-495a-b2bc-830462fe0fa1/visual-basic-console-application-with-arguments?forum=vbgeneral
            'Console.WriteLine("Usage: {0} <source dir> <destination dir> <tar file numbers>", cmd_args(0))
            'https://docs.microsoft.com/zh-tw/dotnet/api/microsoft.visualbasic.applicationservices.assemblyinfo?view=windowsdesktop-6.0
            Console.WriteLine("Version: {0}", My.Application.Info.Version.ToString)
            Console.WriteLine("Usage: {0} <source dir> <destination dir> <tar file numbers> <delete source files> <csv file path>", "tarFiles.exe")
            Exit Sub
        End If

        Dim srcDir As String = cmd_args(1)
        Dim dstDir As String = cmd_args(2)
        Dim fileNums As String = cmd_args(3)
        Dim delSrcFiles As String = cmd_args(4)
        Dim sumCsvFile As String = cmd_args(5)


        Dim num As Integer = 1000
        Dim numDelSrcFiles As Integer = 0

        '檢查輸入
        If Integer.TryParse(fileNums, num) Then
            If num > 2000 Or num < 100 Then
                Console.WriteLine("打包檔案大小請設定介於100-2000數值!!")
                Exit Sub
            Else
                Console.WriteLine("打包檔案設定: {0}", num.ToString)
            End If
        Else
            Console.WriteLine("<tar file numbers>輸入參數非數字")
            Exit Sub
        End If

        If Integer.TryParse(delSrcFiles, numDelSrcFiles) Then
            If numDelSrcFiles = 0 Then
                Console.WriteLine("完成後刪除來源檔案: No")
            ElseIf numDelSrcFiles = 1 Then
                Console.WriteLine("完成後刪除來源檔案: Yes")
            Else
                Console.WriteLine("<delete source files> 請輸入 0 或 1")
            End If
        Else
                Console.WriteLine("<delete source files> 請輸入 0 或 1")
            Exit Sub
        End If


        '準備totalCsv
        Dim sumCsvLogWriter As System.IO.StreamWriter
        sumCsvLogWriter = My.Computer.FileSystem.OpenTextFileWriter(sumCsvFile, True)
        sumCsvLogWriter.WriteLine("原始檔案,目的檔案")


        '建立目的目錄
        If Not Directory.Exists(dstDir) Then
            Directory.CreateDirectory(dstDir)
        End If

        Dim csvLogWriter As System.IO.StreamWriter
        Dim csvLogFile As String = Path.GetFileName(srcDir) + ".csv"

        If Directory.Exists(srcDir) Then
            Dim files() As String = Directory.GetFiles(srcDir, "*.zip", SearchOption.AllDirectories) '僅檢查.zip的檔案
            Console.WriteLine("{0} 檔案總數: {1}", srcDir, files.Count())

            '計算打包檔案數
            Dim packedNum As Integer = Math.Ceiling(files.Count() / num)
            Dim packedFileName As String
            Dim startPos As Integer
            Dim endPos As Integer
            '最後一個打包檔要特別處理
            Dim endPosLast As Integer = files.Count() - 1
            Console.WriteLine(packedNum.ToString)

            '建立csv log 
            csvLogWriter = My.Computer.FileSystem.OpenTextFileWriter(dstDir + "\" + csvLogFile, True)
            csvLogWriter.WriteLine("原始檔案,目的檔案")


            '每個打包檔流程
            For index As Integer = 1 To packedNum
                '計算檔名
                startPos = (index - 1) * num
                endPos = index * num - 1

                '檢查是否是最後一個打包檔
                If files.Count() - endPos <= 0 Then
                    endPos = endPosLast
                End If

                'packedFileName = Path.GetFileName(srcDir) + "_" + index.ToString.PadLeft(12, "0") + ".zip" '當Path是給目錄, GetFileName回傳最後一層目錄名稱
                packedFileName = Path.GetFileName(dstDir) + "_" + index.ToString.PadLeft(12, "0") + ".zip" '當Path是給目錄, GetFileName回傳最後一層目錄名稱
                packedFileName = dstDir + "\" + packedFileName
                Console.WriteLine(packedFileName)
                Console.WriteLine("-----------------------------------------------------------------------------------")

                '產生csv檔案清單
                For i2 As Integer = startPos To endPos
                    'Console.WriteLine(files(i2))
                    csvLogWriter.WriteLine("{0},{1}", files(i2), packedFileName)
                    sumCsvLogWriter.WriteLine("{0},{1}", files(i2), packedFileName)
                Next

                'https://docs.microsoft.com/en-us/dotnet/visual-basic/programming-guide/language-features/arrays/
                Dim filesTmp(endPos - startPos) As String
                Array.Copy(files, startPos, filesTmp, 0, endPos - startPos + 1)
                'For Each aa As String In filesTmp
                '    Console.WriteLine(aa)
                'Next
                tarFile(packedFileName, filesTmp, srcDir)

            Next

            csvLogWriter.Close()

            '檢查是否需刪除來源檔案
            If numDelSrcFiles = 1 Then

                '直接刪目錄 (暫不使用此方式，因files字串陣列僅搜尋*.zip檔案，若直接刪目錄可能會誤刪除未打包檔案)
                'My.Computer.FileSystem.DeleteDirectory(srcDir, FileIO.DeleteDirectoryOption.DeleteAllContents)

                '個別刪除目錄內資料
                'https://docs.microsoft.com/zh-tw/dotnet/visual-basic/developing-apps/programming/drives-directories-files/how-to-delete-a-file
                For Each file As String In files
                    My.Computer.FileSystem.DeleteFile(file, FileIO.UIOption.AllDialogs, FileIO.RecycleOption.SendToRecycleBin)
                    'System.IO.File.Delete(File)
                    Console.WriteLine("Deleting {0}", file)
                Next

            End If

        Else
            Console.WriteLine("來源目錄 [{0}] 不存在!!", srcDir)
            Exit Sub
        End If


        sumCsvLogWriter.Close()





    End Sub

End Module
