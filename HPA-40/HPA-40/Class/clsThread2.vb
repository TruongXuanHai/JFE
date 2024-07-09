' HPA-40
' clsThread2.vb
' Threadのタスクを実行
'
'CORYRIGHT(C) 2024 HAKARU PLUS CORPORATION
'
' 修正履歴
' 2024/06/17 チュオンスアンハイ

Imports System.Globalization
Imports System.IO
Imports System.Threading

Public Class clsThread2
    'Threadsのタスクを実行する
    Public Sub subThread2Processing()
        Dim threadId As Integer = Thread.CurrentThread.ManagedThreadId
        Dim intUnitIndex As Integer = 0
        Dim clsProcessThread2 As New clsThreadProcess

        While Not gblStopFlagGateway
            If Not gblStopTempFlagGateway Then
                If (gblCommStartThr2Flag = True And gblCommEndThr2Flag = False) Then
                    While intUnitIndex <= gw2.UnitNumber - 1
                        Console.WriteLine("Thread2:処理中")
                        clsProcessThread2.subProcessingThread(1, intUnitIndex)
                        intUnitIndex += 1
                        Thread.Sleep(1000)
                    End While

                    If intUnitIndex = gw2.UnitNumber Then
                        'ファイル名を設定
                        Dim formatTime As String = "yyyy/MM/dd HH:mm:ss"
                        Dim dateTimeConvert As DateTime = DateTime.ParseExact(dtTimeTaking, formatTime, CultureInfo.InvariantCulture)
                        Dim strTimePart As String = dateTimeConvert.Year.ToString("0000") & dateTimeConvert.Month.ToString("00") &
                            dateTimeConvert.Day.ToString("00") & dateTimeConvert.Hour.ToString("00") &
                            dateTimeConvert.Minute.ToString("00") & dateTimeConvert.Second.ToString("00")
                        Dim strCSVFilePath As String = strbasePathGateway & gw2.GWName & "_" & strTimePart & ".CSV"
                        Dim strCTCFilePath As String = strbasePathGateway & gw2.GWName & "_" & strTimePart & ".CTC"

                        'ファイルが存在してるかどうか確認
                        If Not File.Exists(strCSVFilePath) OrElse Not File.Exists(strCTCFilePath) Then
                            'フラグを設定
                            gblCommStartThr2Flag = True
                            gblCommEndThr2Flag = False
                            'ユニットのインデックスをリセット
                            intUnitIndex = 0
                        Else
                            'フラグを設定
                            gblCommStartThr2Flag = False
                            gblCommEndThr2Flag = True
                            'ユニットのインデックスをリセット
                            intUnitIndex = 0
                            Console.WriteLine("Thread2の処理が完了しました")
                        End If
                    End If
                End If
            Else
                Thread.Sleep(500)
                Console.WriteLine("Thread2 (ID:{0})の状態: 待機中", threadId)
            End If
        End While

        'While Not gblStopFlagGateway
        '    If Not gblStopTempFlagGateway Then
        '        If (gblCommStartThr2Flag = True And gblCommEndThr2Flag = False) Then
        '            Thread.Sleep(100)
        '            gintDbgCount2GW += 1
        '            Console.WriteLine("Count in Thread2: {0}", gintDbgCount2GW)
        '            Console.WriteLine("Thread2の時間データ: {0}", dtTimeTaking)
        '            '通信開始フラグをFalseにする
        '            gblCommStartThr2Flag = False
        '            '通信終了フラグをTrueにする
        '            gblCommEndThr2Flag = True
        '        End If
        '    Else
        '        Thread.Sleep(500)
        '        Console.WriteLine("Thread2 (ID:{0})の状態: 待機中", threadId)
        '    End If
        'End While
        Console.WriteLine("Thread2 (ID:{0})の状態: 終了", threadId)
    End Sub
End Class
