Attribute VB_Name = "Module1"
Dim 行數 As Long

#If VBA7 Then
 Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
 Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
#End If

Sub 每30秒記錄()
    行數 = 2 ' 開始記錄的行數
    Dim IsConnected1, IsConnected2 As Boolean
    'IsConnected1 = DEBUG_COM_PORT(1)
    IsConnected1 = START_COM_PORT(1, "Baud=9600 Data=8")
    IsConnected2 = START_COM_PORT(9, "Baud=9600 Data=8")
    ' 設定下一次執行的時間
    Application.OnTime Now + TimeValue("00:00:05"), "記錄時間"
    Debug.Print IsConnected1, IsConnected2
End Sub

Sub 記錄時間()
    Dim 目標工作表 As Worksheet
    
    ' 替換「工作表名稱」為實際的工作表名稱
    Set 目標工作表 = ThisWorkbook.Sheets("工作表1")
    
    ' 在目標工作表的下一列寫入當前時間
    目標工作表.Cells(行數, 1).Value = Format(Now(), "hh:mm:ss")
    
    Dim IsSent1, IsSent2 As Boolean
    Dim ReceivedData1, ReceivedData2 As String
    IsSent1 = TRANSMIT_COM_PORT(1, "GETMEAS CH" & Chr(13))
    Sleep 500
    ReceivedData1 = RECEIVE_COM_PORT(1)
    目標工作表.Cells(行數, 2).Value = ReceivedData1
    IsSent2 = TRANSMIT_COM_PORT(9, "GETMEAS CH2 25" & Chr(13))
    Sleep 500
    ReceivedData2 = RECEIVE_COM_PORT(9)
    目標工作表.Cells(行數, 3).Value = ReceivedData2
    
    '增加行數，準備寫入下一列
    行數 = 行數 + 1
    
    ' 設定下一次執行的時間
    Debug.Print Now()
    
    Debug.Print ReceivedData1, ReceivedData2
    Application.OnTime Now + TimeValue("00:00:05"), "記錄時間"
End Sub


