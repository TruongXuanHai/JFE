' HPA-40
' clsThread5.vb
' Threadのタスクを実行
'
'CORYRIGHT(C) 2024 HAKARU PLUS CORPORATION
'
' 修正履歴
' 2024/06/17 チュオンスアンハイ

Imports System.Threading

Public Class clsThread5
    'Threadsのタスクを実行する
    Public Sub subThread5Processing()
        Dim threadId As Integer = Thread.CurrentThread.ManagedThreadId

        While Not gblStopFlagGateway
            If Not gblStopTempFlagGateway Then
                If (gblCommStartThr5Flag = True And gblCommEndThr5Flag = False) Then
                    Thread.Sleep(500)
                    gintDbgCount5GW += 1
                    Console.WriteLine("Count in Thread5: {0}", gintDbgCount5GW)
                    Console.WriteLine("Thread5の時間データ: {0}", dtTimeTaking)
                    '通信開始フラグをFalseにする
                    gblCommStartThr5Flag = False
                    '通信終了フラグをTrueにする
                    gblCommEndThr5Flag = True
                End If
            Else
                Thread.Sleep(500)
                Console.WriteLine("Thread5 (ID:{0})の状態: 待機中", threadId)
            End If
        End While
        Console.WriteLine("Thread5 (ID:{0})の状態: 終了", threadId)
    End Sub
End Class
