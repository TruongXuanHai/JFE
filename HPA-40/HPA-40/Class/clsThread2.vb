' HPA-40
' clsThread2.vb
' Threadのタスクを実行
'
'CORYRIGHT(C) 2024 HAKARU PLUS CORPORATION
'
' 修正履歴
' 2024/06/17 チュオンスアンハイ

Imports System.Threading

Public Class clsThread2
    'Threadsのタスクを実行する
    Public Sub subThread2Processing()
        Dim threadId As Integer = Thread.CurrentThread.ManagedThreadId
        Dim intUnitIndex As Integer = 0
        Dim clsProcessThread2 As New clsThreadProcess

        'While Not gblStopFlagGateway
        '    If Not gblStopTempFlagGateway Then
        '        If (gblCommStartThr2Flag = True And gblCommEndThr2Flag = False) Then
        '            While intUnitIndex <= gw2.UnitNumber - 1
        '                Console.WriteLine("Thread2:処理中")
        '                clsProcessThread2.subProcessingThread(0, intUnitIndex)
        '                intUnitIndex += 1
        '            End While

        '            If intUnitIndex = gw2.UnitNumber Then
        '                '通信開始フラグをFalseにする
        '                gblCommStartThr2Flag = False
        '                '通信終了フラグをTrueにする
        '                gblCommEndThr2Flag = True
        '                'ユニットのインデックスをリセット
        '                intUnitIndex = 0
        '                Console.WriteLine("Thread2の処理が完了しました")
        '            End If
        '        End If
        '    Else
        '        Thread.Sleep(500)
        '        Console.WriteLine("Thread2 (ID:{0})の状態: 待機中", threadId)
        '    End If
        'End While

        While Not gblStopFlagGateway
            If Not gblStopTempFlagGateway Then
                If (gblCommStartThr2Flag = True And gblCommEndThr2Flag = False) Then
                    Thread.Sleep(500)
                    gintDbgCount2GW += 1
                    Console.WriteLine("Count in Thread2: {0}", gintDbgCount2GW)
                    Console.WriteLine("Thread2の時間データ: {0}", dtTimeTaking)
                    '通信開始フラグをFalseにする
                    gblCommStartThr2Flag = False
                    '通信終了フラグをTrueにする
                    gblCommEndThr2Flag = True
                End If
            Else
                Thread.Sleep(500)
                Console.WriteLine("Thread2 (ID:{0})の状態: 待機中", threadId)
            End If
        End While
        Console.WriteLine("Thread2 (ID:{0})の状態: 終了", threadId)
    End Sub
End Class
