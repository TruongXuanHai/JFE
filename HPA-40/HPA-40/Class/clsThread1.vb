' HPA-40
' clsThread1.vb
' Threadのタスクを実行
'
'CORYRIGHT(C) 2024 HAKARU PLUS CORPORATION
'
' 修正履歴
' 2024/06/17 チュオンスアンハイ

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
                    End While

                    If intUnitIndex = gw1.UnitNumber Then
                        '通信開始フラグをFalseにする
                        gblCommStartThr1Flag = False
                        '通信終了フラグをTrueにする
                        gblCommEndThr1Flag = True
                        'ユニットのインデックスをリセット
                        intUnitIndex = 0
                        Console.WriteLine("Thread1の処理が完了しました")
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
