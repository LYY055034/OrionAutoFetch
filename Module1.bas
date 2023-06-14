Attribute VB_Name = "Module1"
Dim ��� As Long

#If VBA7 Then
 Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
 Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
#End If

Sub �C30��O��()
    ��� = 2 ' �}�l�O�������
    Dim IsConnected1, IsConnected2 As Boolean
    'IsConnected1 = DEBUG_COM_PORT(1)
    IsConnected1 = START_COM_PORT(1, "Baud=9600 Data=8")
    IsConnected2 = START_COM_PORT(9, "Baud=9600 Data=8")
    ' �]�w�U�@�����檺�ɶ�
    Application.OnTime Now + TimeValue("00:00:05"), "�O���ɶ�"
    Debug.Print IsConnected1, IsConnected2
End Sub

Sub �O���ɶ�()
    Dim �ؼФu�@�� As Worksheet
    
    ' �����u�u�@��W�١v����ڪ��u�@��W��
    Set �ؼФu�@�� = ThisWorkbook.Sheets("�u�@��1")
    
    ' �b�ؼФu�@���U�@�C�g�J��e�ɶ�
    �ؼФu�@��.Cells(���, 1).Value = Format(Now(), "hh:mm:ss")
    
    Dim IsSent1, IsSent2 As Boolean
    Dim ReceivedData1, ReceivedData2 As String
    IsSent1 = TRANSMIT_COM_PORT(1, "GETMEAS CH" & Chr(13))
    Sleep 500
    ReceivedData1 = RECEIVE_COM_PORT(1)
    �ؼФu�@��.Cells(���, 2).Value = ReceivedData1
    IsSent2 = TRANSMIT_COM_PORT(9, "GETMEAS CH2 25" & Chr(13))
    Sleep 500
    ReceivedData2 = RECEIVE_COM_PORT(9)
    �ؼФu�@��.Cells(���, 3).Value = ReceivedData2
    
    '�W�[��ơA�ǳƼg�J�U�@�C
    ��� = ��� + 1
    
    ' �]�w�U�@�����檺�ɶ�
    Debug.Print Now()
    
    Debug.Print ReceivedData1, ReceivedData2
    Application.OnTime Now + TimeValue("00:00:05"), "�O���ɶ�"
End Sub


