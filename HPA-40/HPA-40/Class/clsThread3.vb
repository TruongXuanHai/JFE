' HPA-40
' clsThread3.vb
' Threadのタスクを実行
'
'CORYRIGHT(C) 2024 HAKARU PLUS CORPORATION
'
' 修正履歴
' 2024/06/17 チュオンスアンハイ

Imports System.Threading

Public Class clsThread3
    'Threadsのタスクを実行する
    Public Sub subThread3Processing()
        Dim threadId As Integer = Thread.CurrentThread.ManagedThreadId

        While Not gblStopFlagGateway
            If Not gblStopTempFlagGateway Then
                If (gblCommStartThr3Flag = True And gblCommEndThr3Flag = False) Then
                    Thread.Sleep(500)
                    gintDbgCount3GW += 1
                    Console.WriteLine("Count in Thread3: {0}", gintDbgCount3GW)
                    Console.WriteLine("Thread3の時間データ: {0}", dtTimeTaking)
                    '通信開始フラグをFalseにする
                    gblCommStartThr3Flag = False
                    '通信終了フラグをTrueにする
                    gblCommEndThr3Flag = True
                End If
            Else
                Thread.Sleep(500)
                Console.WriteLine("Thread3 (ID:{0})の状態: 待機中", threadId)
            End If
        End While
        Console.WriteLine("Thread3 (ID:{0})の状態: 終了", threadId)
    End Sub
End Class
