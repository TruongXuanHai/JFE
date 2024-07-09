' HPA-40
' clsThread1.vb
' Threadのタスクを実行
'
'CORYRIGHT(C) 2024 HAKARU PLUS CORPORATION
'
' 修正履歴
' 2024/06/17 チュオンスアンハイ

Imports System.Globalization
Imports System.IO
Imports System.Threading
Public Class clsThread1
    'Threadsのタスクを実行する
    Public Sub subThread1Processing()
        Dim threadId As Integer = Thread.CurrentThread.ManagedThreadId
        Dim intUnitIndex As Integer = 0
        Dim clsProcessThread1 As New clsThreadProcess

        While Not gblStopFlagGateway
            If Not gblStopTempFlagGateway Then
                If (gblCommStartThr1Flag = True And gblCommEndThr1Flag = False) Then
                    While intUnitIndex <= gw1.UnitNumber - 1
                        Console.WriteLine("Thread1:処理中")
                        clsProcessThread1.subProcessingThread(0, intUnitIndex)
                        intUnitIndex += 1
                        Thread.Sleep(1000)
                    End While

                    If intUnitIndex = gw1.UnitNumber Then
                        'ファイル名を設定
                        Dim formatTime As String = "yyyy/MM/dd HH:mm:ss"
                        Dim dateTimeConvert As DateTime = DateTime.ParseExact(dtTimeTaking, formatTime, CultureInfo.InvariantCulture)
                        Dim strTimePart As String = dateTimeConvert.Year.ToString("0000") & dateTimeConvert.Month.ToString("00") &
                            dateTimeConvert.Day.ToString("00") & dateTimeConvert.Hour.ToString("00") &
                            dateTimeConvert.Minute.ToString("00") & dateTimeConvert.Second.ToString("00")
                        Dim strCSVFilePath As String = strbasePathGateway & gw1.GWName & "_" & strTimePart & ".CSV"
                        Dim strCTCFilePath As String = strbasePathGateway & gw1.GWName & "_" & strTimePart & ".CTC"

                        'ファイルが存在してるかどうか確認
                        If Not File.Exists(strCSVFilePath) OrElse Not File.Exists(strCTCFilePath) Then
                            'フラグを設定
                            gblCommStartThr1Flag = True
                            gblCommEndThr1Flag = False
                            'ユニットのインデックスをリセット
                            intUnitIndex = 0
                        Else
                            '通信開始フラグをFalseにする
                            gblCommStartThr1Flag = False
                            '通信終了フラグをTrueにする
                            gblCommEndThr1Flag = True
                            'ユニットのインデックスをリセット
                            intUnitIndex = 0
                            Console.WriteLine("Thread1の処理が完了しました")
                        End If
                    End If

                End If
            Else
                Thread.Sleep(500)
                Console.WriteLine("Thread1 (ID:{0})の状態: 待機中", threadId)
            End If
        End While
        Console.WriteLine("Thread1 (ID:{0})の状態: 終了", threadId)
    End Sub
End Class
