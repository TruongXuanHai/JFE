' HPA-40
' getXmlData.vb
' XMLファイルから設定値を取得
'
'CORYRIGHT(C) 2024 HAKARU PLUS CORPORATION
'
' 修正履歴
' 2024/06/17 チュオンスアンハイ

Imports System.IO
Imports System.Net
Imports System.Threading
Imports System.Threading.Tasks

Public Class clsTaskServer
#Region "ディレクトリ"
    Private strbasePathGateway As String = funcGetAppPath() & "\Gateway\"
    Private strbasePathServer As String = funcGetAppPath() & "\Server\"
    Private strbasePathFailed As String = funcGetAppPath() & "\Failed\"
#End Region

#Region "変数"
    Private sttFileCSVTranfer As Boolean = False
    Private sttFileCTCTranfer As Boolean = False
#End Region

#Region "関数_本アプリの起動ディレクトリパスを取得する"
    '本アプリの起動ディレクトリパスを取得する
    Private Function funcGetAppPath() As String
        Dim objFileInfo As New System.IO.FileInfo(System.Reflection.Assembly.GetExecutingAssembly().Location)
        Return objFileInfo.DirectoryName
    End Function
#End Region

#Region "関数_Task2を作成"
    Public Sub subInitializeTaskServer()
        'タスクの存在を確認
        If gtaskServerProcessing Is Nothing Then
            gtaskServerProcessing = New Task(Sub() subRunTaskServer())
            gtaskServerProcessing.Start()
        End If
    End Sub
#End Region

#Region "関数_Task2を実行"
    Private Sub subRunTaskServer()
        While Not gblStopFlagServer
            If Not gblStopTempFlagServer Then
                'すべてのCTCファイルをリスト化する
                Dim arrCTCFiles As String() = Directory.GetFiles(strbasePathGateway, "*.CTC")

                'CTCファイルがある
                If arrCTCFiles.Length > 0 Then
                    For Each strCTCFile As String In arrCTCFiles
                        '転送状態を設定
                        sttFileCSVTranfer = False
                        sttFileCTCTranfer = False

                        'CSVファイル名を設定
                        Dim strCSVFileCheck As String = Path.ChangeExtension(strCTCFile, ".CSV")

                        'CSVファイルが存在している
                        If File.Exists(strCSVFileCheck) Then
                            'サーバCSVファイルを転送　→ステータスに戻る
                            sttFileCSVTranfer = funcUploadFileToFtp(ipAddressInfo, userNameInfo, passWordInfo, pathInfo, strCSVFileCheck)

                            'サーバCTCファイルを転送　→ステータスに戻る
                            sttFileCTCTranfer = funcUploadFileToFtp(ipAddressInfo, userNameInfo, passWordInfo, pathInfo, strCTCFile)

                            '転送が成功になる際ファイルの移動を実行
                            If sttFileCSVTranfer And sttFileCTCTranfer Then
                                'ServerフォルダにCSV/CTCファイルを移動
                                subMoveFileToServerFolder(strCSVFileCheck, strCTCFile)
                            Else
                                'FailedフォルダにCSV/CTCファイルを移動
                                subMoveFileToFailed(strCSVFileCheck, strCTCFile)
                            End If
                        End If
                    Next
                Else
                    Thread.Sleep(500)
                End If
            Else
                Thread.Sleep(500)
                'Console.WriteLine("Task2の状態: 待機中")
            End If
        End While
        Console.WriteLine("Task2の状態: 終了")
    End Sub
#End Region

#Region "関数_サーバにファイルアップロード"
    Private Function funcUploadFileToFtp(ByVal ftpServer As String, ByVal ftpUsername As String, ByVal ftpPassword As String, ByVal newFolderName As String, ByVal localFilePath As String) As Boolean
        Try
            'ファイル名を取る
            Dim strfileName As String = Path.GetFileName(localFilePath)
            Dim struploadUrl As String = "ftp://" & ftpServer.TrimEnd("/") & "/" & newFolderName & "/" & strfileName

            'FTPリクエストを作成
            Dim request As FtpWebRequest = DirectCast(WebRequest.Create(struploadUrl), FtpWebRequest)
            request.Method = WebRequestMethods.Ftp.UploadFile
            request.Credentials = New NetworkCredential(ftpUsername, ftpPassword)

            'ファイルをロード
            Dim fileContents() As Byte = File.ReadAllBytes(localFilePath)
            request.ContentLength = fileContents.Length

            ' ファイルをリクエストストリームに書き込む　(転送)
            Using requestStream As Stream = request.GetRequestStream()
                requestStream.Write(fileContents, 0, fileContents.Length)
            End Using

            'リスポーンを取る
            Using response As FtpWebResponse = DirectCast(request.GetResponse(), FtpWebResponse)
                'サーバへ転送することが成功になる
                If response.StatusDescription.Contains("226 Transfer complete") Then
                    Console.WriteLine("{0}をアップロードできました。ステータス: {1}", strfileName, response.StatusDescription)
                    Return True

                    'サーバへ転送することが失敗になる
                Else
                    Console.WriteLine("{0}をアップロードできませんでした。ステータス: {1}", strfileName, response.StatusDescription)
                    Return False
                End If
            End Using

        Catch ex As Exception
            ' エラーメッセージを表示する
            Console.WriteLine("ファイルをアップロードできませんでした。エラー: {0}", ex.Message)
            Return False
        End Try
    End Function
#End Region

#Region "ファイルを移動"
    Private Sub subMoveFileToFailed(ByVal strSourceFileCSVPath As String, ByVal strSourceFileCTCPath As String)
        Dim strDestFileCSVPath As String = Path.Combine(strbasePathFailed, Path.GetFileName(strSourceFileCSVPath))
        Dim strDestFileCTCPath As String = Path.Combine(strbasePathFailed, Path.GetFileName(strSourceFileCTCPath))

        'ファイルを移動
        File.Move(strSourceFileCSVPath, strDestFileCSVPath)
        File.Move(strSourceFileCTCPath, strDestFileCTCPath)
    End Sub
#End Region

    Private Sub subMoveFileToServerFolder(ByVal strSourceFileCSVPath As String, ByVal strSourceFileCTCPath As String)
        Dim strFileCSVName As String = Path.GetFileName(strSourceFileCSVPath)
        Dim strFileCTCName As String = Path.GetFileName(strSourceFileCTCPath)
        Dim strDatePart As String = Mid(strFileCSVName, InStrRev(strFileCSVName, "_") + 1, 8)
        Dim strYearFolder As String = Path.Combine(strbasePathServer, Mid(strDatePart, 1, 4))
        Dim strMonthFolder As String = Path.Combine(strYearFolder, Mid(strDatePart, 5, 2))

        '年フォルダを作成
        If Not Directory.Exists(strYearFolder) Then
            Directory.CreateDirectory(strYearFolder)
        End If
        '月フォルダを作成
        If Not Directory.Exists(strMonthFolder) Then
            Directory.CreateDirectory(strMonthFolder)
        End If

        Dim destFileCSVPath As String = Path.Combine(strMonthFolder, strFileCSVName)
        Dim destFileCTCPath As String = Path.Combine(strMonthFolder, strFileCTCName)

        'CSV/CTCファイルを移動
        If File.Exists(destFileCSVPath) Then
            File.Delete(destFileCSVPath)
        End If
        File.Move(strSourceFileCSVPath, destFileCSVPath)

        If File.Exists(destFileCTCPath) Then
            File.Delete(destFileCTCPath)
        End If
        File.Move(strSourceFileCTCPath, destFileCTCPath)
    End Sub

End Class
