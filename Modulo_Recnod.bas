Attribute VB_Name = "Modulo_Recnod"
'*******************************************************************************
'
'*******************************************************************************

Private Const TH32CS_SNAPPROCESS As Long = 2&
Private Const MAX_PATH As Integer = 260
Private Const INFINITE = &HFFFF      '  Infinite timeout
Private Const PROCESS_TERMINATE = &H1
Private Const FLASHW_CAPTION = &H1

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" _
   (ByVal lFlags As Long, ByVal lProcessID As Long) As Long

Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" _
   (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long

Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" _
   (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long

Private Declare Sub CloseHandle Lib "kernel32" _
   (ByVal hPass As Long)

Private Declare Function OpenProcess Lib "kernel32" _
   (ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long

Private Declare Function TerminateProcess Lib "kernel32" _
   (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
   
Private Declare Function WaitForSingleObject Lib "kernel32" _
   (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
'*******************************************************************************
' Method - Obtem_IdProcesso
'*******************************************************************************
Private Function Obtem_IdProcesso(ByVal strProcesso As String) As Long
    
    Dim hSnapShot As Long
    Dim uProcess As PROCESSENTRY32
    Dim r As Long
    Dim I As Long
    Dim posBar As Long
    Dim idProcess As Long
    
    strProcesso = UCase(strProcesso)
    Obtem_IdProcesso = 0
    
    '***************************************************************************
    ' Pesquisa todos os processos ativos
    '***************************************************************************
    
    hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    If hSnapShot = 0 Then Exit Function

    uProcess.dwSize = Len(uProcess)
    r = ProcessFirst(hSnapShot, uProcess)

    Do While r
        For I = 1 To Len(uProcess.szExeFile)
            If Mid(uProcess.szExeFile, I, 1) = "\" Then
                posBar = I
            End If
        Next
        ' Obtem o número de identificação do processo
        If UCase(Mid(uProcess.szExeFile, posBar + 1, InStr(uProcess.szExeFile, Chr(0)) - 1)) = strProcesso Then
            idProcess = uProcess.th32ProcessID
        End If

        r = ProcessNext(hSnapShot, uProcess)

    Loop

    Call CloseHandle(hSnapShot)
    Obtem_IdProcesso = idProcess
    
End Function

'*******************************************************************************
' Method - Termina_IdProcesso
'*******************************************************************************
Private Function Termina_IdProcesso(ByVal lngProcesso As Long) As Boolean
    
    '***************************************************************************
    ' Termina o processo
    '***************************************************************************
    
    Dim hProcess As Long
    Dim bRet As Long
    Dim stat As Long
    
    Termina_IdProcesso = False
    
    '***************************************************************************
    ' Obtem um handler para o processo. Se verdadeiro, espera
    ' que o processo termine antes de retornar.
    '***************************************************************************
    
    hProcess = OpenProcess(PROCESS_TERMINATE, 0&, lngProcesso)
    If hProcess <> 0 Then

        bRet = TerminateProcess(hProcess, &HFFFFFFFF)
        If bRet Then
            Call WaitForSingleObject(hProcess, INFINITE)
        End If


        Call CloseHandle(hProcess)
        Termina_IdProcesso = True
        
    End If

End Function
'*******************************************************************************
' Method - Matador
'*******************************************************************************
Public Sub Matador()
    
    Dim Flag As Long
    Dim Inicio As Date
    Dim Fim As Date
    
    Flag = Obtem_IdProcesso("RecNod.exe")
    DoEvents
    
    Inicio = Now
    While (Flag <> 0)
        Fim = Now
        If (DateDiff("s", Inicio, Fim) > 3) Then
            Flag = 0
        Else
            Flag = Obtem_IdProcesso("RecNod.exe")
        End If
        DoEvents
    Wend
    
    Flag = Obtem_IdProcesso("RecNod.exe")
    DoEvents

    While (Flag <> 0)
        Termina_IdProcesso (Flag)
        DoEvents
        
        Flag = Obtem_IdProcesso("RecNod.exe")
        DoEvents
    Wend
    
End Sub
'*******************************************************************************
' Method - Avalia_Aptidao
'*******************************************************************************
'Public Function Avalia_Aptidao(Num_Formiga As Integer, Num_Elementos As Integer, Lim_Esp_Tec As Double) As Double
Public Sub Avalia_Aptidao()
        
    Dim X_Y_Elemento(1 To 20, 1 To 3) As Integer
    Dim M As Integer
    Dim Maior_Pico As Double
    Dim Boro As Double
    Dim Padrao(1 To 20, 1 To 3) As Integer
    Dim Linha As Integer
    Dim Pico As Double
    Dim Linha_Pico As String
    Dim Linha_Boro As String
    Dim Lim_Esp_Tec As Double
    Dim Kuato(1 To 20) As Integer
    Dim Avalia_Aptidao As Double
    
    '************************************************************************
    ' *** Le arquivo gerado pelo Algoritmo ***
    '************************************************************************
    
    On Error GoTo AVALIA_APTIDAO_ERRO_001
    
    Open App.Path + "/Matheus_Saida.dat" For Input As #1
    For Linha = 1 To 20
        Input #1, Kuato(Linha)
    Next Linha
    Close #1
    
    '************************************************************************
    '  Original Chapot - MAPA 1
    '************************************************************************
    
    X_Y_Elemento(1, 1) = 7: X_Y_Elemento(1, 2) = 1: X_Y_Elemento(1, 3) = 2
    X_Y_Elemento(2, 1) = 5: X_Y_Elemento(2, 2) = 1: X_Y_Elemento(2, 3) = 4
    X_Y_Elemento(3, 1) = 6: X_Y_Elemento(3, 2) = 1: X_Y_Elemento(3, 3) = 2
    X_Y_Elemento(4, 1) = 3: X_Y_Elemento(4, 2) = 1: X_Y_Elemento(4, 3) = 4
    X_Y_Elemento(5, 1) = 2: X_Y_Elemento(5, 2) = 1: X_Y_Elemento(5, 3) = 6
    X_Y_Elemento(6, 1) = 4: X_Y_Elemento(6, 2) = 1: X_Y_Elemento(6, 3) = 4
    
    X_Y_Elemento(7, 1) = 4: X_Y_Elemento(7, 2) = 4: X_Y_Elemento(7, 3) = 1
    X_Y_Elemento(8, 1) = 5: X_Y_Elemento(8, 2) = 5: X_Y_Elemento(8, 3) = 2
    X_Y_Elemento(9, 1) = 3: X_Y_Elemento(9, 2) = 3: X_Y_Elemento(9, 3) = 4
    X_Y_Elemento(10, 1) = 2: X_Y_Elemento(10, 2) = 2: X_Y_Elemento(10, 3) = 5
    
    X_Y_Elemento(11, 1) = 6: X_Y_Elemento(11, 2) = 2: X_Y_Elemento(11, 3) = 3
    X_Y_Elemento(12, 1) = 5: X_Y_Elemento(12, 2) = 2: X_Y_Elemento(12, 3) = 4
    X_Y_Elemento(13, 1) = 4: X_Y_Elemento(13, 2) = 2: X_Y_Elemento(13, 3) = 4
    X_Y_Elemento(14, 1) = 5: X_Y_Elemento(14, 2) = 3: X_Y_Elemento(14, 3) = 6
    X_Y_Elemento(15, 1) = 7: X_Y_Elemento(15, 2) = 2: X_Y_Elemento(15, 3) = 5
    X_Y_Elemento(16, 1) = 6: X_Y_Elemento(16, 2) = 3: X_Y_Elemento(16, 3) = 3
    X_Y_Elemento(17, 1) = 6: X_Y_Elemento(17, 2) = 4: X_Y_Elemento(17, 3) = 2
    X_Y_Elemento(18, 1) = 5: X_Y_Elemento(18, 2) = 4: X_Y_Elemento(18, 3) = 5
    X_Y_Elemento(19, 1) = 4: X_Y_Elemento(19, 2) = 3: X_Y_Elemento(19, 3) = 6
    X_Y_Elemento(20, 1) = 3: X_Y_Elemento(20, 2) = 2: X_Y_Elemento(20, 3) = 4
       
    '************************************************************************
    ' *** Enche a matriz padrão com as soluções  ***
    '************************************************************************
    'For M = 1 To Num_Elementos
    '************************************************************************
    
    For M = 1 To 20
        'Padrao(M, 1) = X_Y_Elemento(Formiga(Num_Formiga).Rota(M), 1)
        'Padrao(M, 2) = X_Y_Elemento(Formiga(Num_Formiga).Rota(M), 2)
        'Padrao(M, 3) = X_Y_Elemento(Formiga(Num_Formiga).Rota(M), 3)
        Padrao(M, 1) = X_Y_Elemento(Kuato(M), 1)
        Padrao(M, 2) = X_Y_Elemento(Kuato(M), 2)
        Padrao(M, 3) = X_Y_Elemento(Kuato(M), 3)
    Next M
        
    '************************************************************************
    ' *** Escreve arquivo Inpnod ***
    '************************************************************************
    
    On Error GoTo AVALIA_APTIDAO_ERRO_002
    
    Open App.Path & "/Minpnod.dat" For Output As #3
    
    Print #3, "6   8   8  15  15"
    Print #3, "1   2   2   2   2   2   2   2"
    Print #3, "9.9098  19.8196 19.8196  19.8196  19.8196  19.8196  19.8196  19.8196"
    Print #3, "1   2   2   2   2   2   2   2"
    Print #3, "9.9098  19.8196 19.8196  19.8196  19.8196  19.8196  19.8196  19.8196"
    Print #3, "1.0E-06    1.0E-03    1.45       1     3000     0.5   37  1  225"
    Print #3, "2          151"
    Print #3, "1.0        0.0"
    Print #3, "2.261E-04  2.207E-04  1.80E-04   2.263E-04  2.263E-04  1.986E-04  1.793E-04  2.263E-04"
    Print #3, "2.263E-04  1.13E-04   1.13E-04   2.264E-04  2.261E-04  1.13E-04   1.13E-04"
    Print #3, "2.207E-04  2.266E-04  2.286E-04  2.270E-04  2.256E-04  2.260E-04  2.265E-04  2.264E-04"
    Print #3, "2.263E-04  1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04"
    Print #3, "1.80E-04   2.286E-04  2.294E-04  2.291E-04  2.286E-04  2.264E-04  2.267E-04  2.265E-04"
    Print #3, "2.264E-04  1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04"
    Print #3, "2.263E-04  2.270E-04  2.291E-04  1.763E-04  1.862E-04  2.273E-04  2.045E-04  1.827E-04"
    Print #3, "2.136E-04  1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04"
    Print #3, "2.263E-04  2.256E-04  2.286E-04  1.862E-04  2.218E-04  2.259E-04  1.876E-04  1.718E-04"
    Print #3, "1.80E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04"
    Print #3, "1.986E-04  2.260E-04  2.264E-04  2.273E-04  2.259E-04  2.263E-04  2.265E-04  1.13E-04"
    Print #3, "1.13E-04   2.266E-04  2.265E-04  1.13E-04   1.13E-04   1.13E-04   1.13E-04"
    Print #3, "1.793E-04  2.265E-04  2.267E-04  2.045E-04  1.876E-04  2.265E-04  2.267E-04  1.13E-04"
    Print #3, "1.13E-04   2.266E-04  2.266E-04  1.13E-04   1.13E-04   1.13E-04   1.13E-04"
    Print #3, "2.263E-04  2.264E-04  2.265E-04  1.827E-04  1.718E-04  1.13E-04   1.13E-04   1.13E-04"
    Print #3, "1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04"
    Print #3, "2.263E-04  2.263E-04  2.264E-04  2.136E-04  1.80E-04   1.13E-04   1.13E-04   1.13E-04"
    Print #3, "1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04"
    Print #3, "1.13E-04   1.13E-04   1.13E-04   1.13E-04   2.266E-04  2.266E-04  1.13E-04   1.13E-04"
    Print #3, "1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04"
    Print #3, "1.13E-04   1.13E-04   1.13E-04   1.13E-04   2.265E-04  2.266E-04  1.13E-04   1.13E-04"
    Print #3, "1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04"
    Print #3, "2.264E-04  1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04"
    Print #3, "1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04"
    Print #3, "2.261E-04  1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04"
    Print #3, "1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04"
    Print #3, "1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04"
    Print #3, "1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04"
    Print #3, "1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04"
    Print #3, "1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04   1.13E-04"
    Print #3, "12861   0       0       13007   13004   13126   12897   13068   13022   9857"
    Print #3, "5906    11655   7550    0       0"
    Print #3, "0       0       0       13253   13227   13158   13067   12823   13165   13627"
    Print #3, "12481   0       0       0       0"
    Print #3, "0       0       0       13301   13252   13203   13111   13116   13355   15272"
    Print #3, "14892   0       0       0       0"
    Print #3, "13007   13253   13301   13042   13180   0       0       0       0       13808"
    Print #3, "10359   0       0       0       0"
    Print #3, "13004   13227   13252   13180   13323   0       0       0       0       12690"
    Print #3, "8758    0       0       0       0"
    Print #3, "13126   13158   13203   0       0       13294   14913   0       0       8274"
    Print #3, "11090   0       0       0       0"
    Print #3, "12897   13067   13111   0       0       14913   15481   0       0       4907"
    Print #3, "7219    0       0       0       0"
    Print #3, "13068   12823   13116   0       0       0       0       5504    8621    0"
    Print #3, "0       0       0       0       0"
    Print #3, "13022   13165   13355   0       0       0       0       8621    11743   0"
    Print #3, "0       0       0       0       0"
    Print #3, "9857    13627   15272   13808   12690   8274    4907    0       0       0"
    Print #3, "0       0       0       0       0"
    Print #3, "5906    12481   14892   10359   8758    11090   7219    0       0       0"
    Print #3, "0       0       0       0       0"
    Print #3, "11655   0       0       0       0       0       0       0       0       0"
    Print #3, "0       0       0       0       0"
    Print #3, "7550    0       0       0       0       0       0       0       0       0"
    Print #3, "0       0       0       0       0"
    Print #3, "0       0       0       0       0       0       0       0       0       0"
    Print #3, "0       0       0       0       0"
    Print #3, "0       0       0       0       0       0       0       0       0       0"
    Print #3, "0       0       0       0       0"
    Print #3, "107.9097        1.0     1876.0  0.5"
    Print #3, "2.9239E-05      2.11816E-05     3.626E-06"
    Print #3, "0.0     237.08  553.5   790.7   1581.3  1581.38   1581.32   1581.4   1510.2"
    Print #3, "0.0     144     335.49  479.27  958.47  958.52    958.49    958.53   915.38"
    Print #3, "3.134E-08  0.0        0.0        3.156E-08  3.150E-08  3.174E-08  3.143E-08"
    Print #3, "3.130E-08  3.113E-08  3.426E-08  2.772E-08  3.702E-08  3.050E-08  0.0       0.0"
    Print #3, "0.0        0.0        0.0        3.184E-08  3.185E-08  3.147E-08  3.145E-08"
    Print #3, "3.069E-08  3.126E-08  3.991E-08  3.829E-08  0.0        0.0        0.0       0.0"
    Print #3, "0.0        0.0        0.0        3.180E-08  3.181E-08  3.139E-08  3.140E-08"
    Print #3, "3.117E-08  3.161E-08  4.217E-08  4.156E-08  0.0        0.0        0.0       0.0"
    Print #3, "3.156E-08  3.184E-08  3.180E-08  3.113E-08  3.150E-08  0.0        0.0"
    Print #3, "0.0        0.0        4.003E-08  3.532E-08  0.0        0.0        0.0       0.0"
    Print #3, "3.150E-08  3.185E-08  3.181E-08  3.150E-08  3.183E-08  0.0        0.0"
    Print #3, "0.0        0.0        3.852E-08  3.262E-08  0.0        0.0        0.0       0.0"
    Print #3, "3.174E-08  3.147E-08  3.139E-08  0.0        0.0        4.010E-08  4.253E-08"
    Print #3, "0.0        0.0        3.162E-08  3.610E-08  0.0        0.0        0.0       0.0"
    Print #3, "3.143E-08  3.145E-08  3.140E-08  0.0        0.0        4.253E-08  4.357E-08"
    Print #3, "0.0        0.0        2.598E-08  2.983E-08  0.0        0.0        0.0       0.0"
    Print #3, "3.130E-08  3.069E-08  3.117E-08  0.0        0.0        0.0        0.0"
    Print #3, "2.693E-08  3.219E-08  0.0        0.0        0.0        0.0        0.0       0.0"
    Print #3, "3.113E-08  3.126E-08  3.161E-08  0.0        0.0        0.0        0.0"
    Print #3, "3.219E-08  3.693E-08  0.0        0.0        0.0        0.0        0.0       0.0"
    Print #3, "3.426E-08  3.991E-08  4.217E-08  4.003E-08  3.852E-08  3.162E-08  2.598E-08"
    Print #3, "0.0        0.0        0.0        0.0        0.0        0.0        0.0       0.0"
    Print #3, "2.772E-08  3.829E-08  4.156E-08  3.532E-08  3.262E-08  3.610E-08  2.983E-08"
    Print #3, "0.0        0.0        0.0        0.0        0.0        0.0        0.0       0.0"
    Print #3, "3.702E-08  0.0        0.0        0.0        0.0        0.0        0.0"
    Print #3, "0.0        0.0        0.0        0.0        0.0        0.0        0.0       0.0"
    Print #3, "3.050E-08  0.0        0.0        0.0        0.0        0.0        0.0"
    Print #3, "0.0        0.0        0.0        0.0        0.0        0.0        0.0       0.0"
    Print #3, "0.0        0.0        0.0        0.0        0.0        0.0        0.0"
    Print #3, "0.0        0.0        0.0        0.0        0.0        0.0        0.0       0.0"
    Print #3, "0.0        0.0        0.0        0.0        0.0        0.0        0.0"
    Print #3, "0.0        0.0        0.0        0.0        0.0        0.0        0.0       0.0"
    Print #3, "2000.0   1.0E-06"
          
    Print #3, " 1  1  1  1  4"
    Print #3, 2; 1; Padrao(1, 1); Padrao(1, 2); Padrao(1, 3)
    Print #3, 3; 1; Padrao(2, 1); Padrao(2, 2); Padrao(2, 3)
    Print #3, 4; 1; Padrao(3, 1); Padrao(3, 2); Padrao(3, 3)
    Print #3, 5; 1; Padrao(4, 1); Padrao(4, 2); Padrao(4, 3)
    Print #3, 6; 1; Padrao(5, 1); Padrao(5, 2); Padrao(5, 3)
    Print #3, 7; 1; Padrao(6, 1); Padrao(6, 2); Padrao(6, 3)
    Print #3, 3; 2; Padrao(11, 1); Padrao(11, 2); Padrao(11, 3)
    Print #3, 4; 2; Padrao(12, 1); Padrao(12, 2); Padrao(12, 3)
    Print #3, 5; 2; Padrao(13, 1); Padrao(13, 2); Padrao(13, 3)
    Print #3, 6; 2; Padrao(14, 1); Padrao(14, 2); Padrao(14, 3)
    Print #3, 7; 2; Padrao(15, 1); Padrao(15, 2); Padrao(15, 3)
    Print #3, 4; 3; Padrao(16, 1); Padrao(16, 2); Padrao(16, 3)
    Print #3, 5; 3; Padrao(17, 1); Padrao(17, 2); Padrao(17, 3)
    Print #3, 6; 3; Padrao(18, 1); Padrao(18, 2); Padrao(18, 3)
    Print #3, 5; 4; Padrao(19, 1); Padrao(19, 2); Padrao(19, 3)
    Print #3, 6; 4; Padrao(20, 1); Padrao(20, 2); Padrao(20, 3)
    Print #3, 2; 2; Padrao(7, 1); Padrao(7, 2); Padrao(7, 3)
    Print #3, 3; 3; Padrao(8, 1); Padrao(8, 2); Padrao(8, 3)
    Print #3, 4; 4; Padrao(9, 1); Padrao(9, 2); Padrao(9, 3)
    Print #3, 5; 5; Padrao(10, 1); Padrao(10, 2); Padrao(10, 3)
    
    Close #3

    '************************************************************************
    '   *** Avalia a aptidão ***
    '************************************************************************
    '   *** Inicializa os arquivos angrab.end e angrap.end ***
    '   *** Tentativa de consertar o bug do Recnod ***
    '************************************************************************
    
    On Error GoTo AVALIA_APTIDAO_ERRO_003
    
    Open App.Path & "/angrab.end" For Output As #7
    Print #7, 10
    Close #7
   
    On Error GoTo AVALIA_APTIDAO_ERRO_004
   
    Open App.Path & "/angrap.end" For Output As #8
    For M = 1 To 64
        Print #8, 10
    Next M
    Close #8
   
    '************************************************************************
    '   *** Executa o RecNod ***
    '************************************************************************

    Shell App.Path & "/recnod.exe", vbMinimizedNoFocus
    Call Matador

    '************************************************************************
    '   *** Lê saída Pico ***
    '************************************************************************
    
    On Error GoTo AVALIA_APTIDAO_ERRO_005
    
    Maior_Pico = -1
    Open App.Path & "/angrap.end" For Input As #4
    
    For Linha = 1 To 64
        Line Input #4, Linha_Pico
        'Linha_Pico = Replace(Linha_Pico, ".", ",")
        Pico = CDbl(Trim(Linha_Pico))
        If Pico > Maior_Pico Then
            Maior_Pico = Pico
        End If
    Next Linha
    Close #4

    '************************************************************************
    '   *** Lê saída Boro
    '************************************************************************

    On Error GoTo AVALIA_APTIDAO_ERRO_006
    Open App.Path & "\angrab.end" For Input As #5
   
    Line Input #5, Linha_Boro
    Boro = CDbl(Trim(Linha_Boro))
    Close #5
                               
    '************************************************************************
    ' *** Função aptidão ***
    '************************************************************************
    
    Lim_Esp_Tec = 1.395
    Avalia_Aptidao = 1E+300
    
    If Maior_Pico > Lim_Esp_Tec Then
        Avalia_Aptidao = Maior_Pico
    Else
        Avalia_Aptidao = 1 / Boro
    End If
    
    '************************************************************************
    ' *** Gravar arquivo aptidão ***
    '************************************************************************
    
    On Error GoTo AVALIA_APTIDAO_ERRO_007
    
    Open App.Path + "/Matheus_Pico_Boro.dat" For Output As #1
    Print #1, Format(Maior_Pico, "00.000000")
    Print #1, Format(Boro, "00.000000")
    Close #1
    
    End
    
    '************************************************************************
    ' *** Tratamento de Erros ***
    '************************************************************************
    
AVALIA_APTIDAO_ERRO_001:
    MsgErro_Avalia_Aptidao ("AVALIA_APTIDAO_ERRO_001")
    End

AVALIA_APTIDAO_ERRO_002:
    MsgErro_Avalia_Aptidao ("AVALIA_APTIDAO_ERRO_002")
    End

AVALIA_APTIDAO_ERRO_003:
    MsgErro_Avalia_Aptidao ("AVALIA_APTIDAO_ERRO_003")
    End

AVALIA_APTIDAO_ERRO_004:
    MsgErro_Avalia_Aptidao ("AVALIA_APTIDAO_ERRO_004")
    End

AVALIA_APTIDAO_ERRO_005:
    MsgErro_Avalia_Aptidao ("AVALIA_APTIDAO_ERRO_005")
    End

AVALIA_APTIDAO_ERRO_006:
    MsgErro_Avalia_Aptidao ("AVALIA_APTIDAO_ERRO_006")
    End

AVALIA_APTIDAO_ERRO_007:
    MsgErro_Avalia_Aptidao ("AVALIA_APTIDAO_ERRO_007")
    End

End Sub
'*******************************************************************************
' Method - Avalia_Aptidao
'*******************************************************************************
Public Sub MsgErro_Avalia_Aptidao(Mensagem)
    
    Prompt = Now & " - " & Mensagem & " - " & Err.Description
    'MsgBox Prompt, vbOK, "Formiga Veneno - Rotina - Avalia_Aptidao"

    Open App.Path + "/Matheus.log" For Append As #9
    Print #9, Prompt
    Close #9

    Dim Aptidao As Double
    Dim Pico As Double
    Dim Boro  As Double

    Aptidao = 1E+300
    Pico = 10#
    Boro = 10#
    
    Open App.Path + "/Matheus_Aptidao.dat" For Output As #1
    Print #1, Pico
    Print #1, Boro
    Close #1


End Sub
'*******************************************************************************


