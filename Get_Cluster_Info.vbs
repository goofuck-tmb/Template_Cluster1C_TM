Option Explicit

' Зададим параметры доступа к кластеру 1С
Dim ServerName                ' имя кластера 1С:Предприятия
Dim ClusterAdmin            ' Имя админа кластера, если есть
Dim ClusterAdminPassword    ' Пароль админа кластера, если есть.

ServerName = "servername"
ClusterAdmin = ""
ClusterAdminPassword = ""
const CLcount = 1 'колличество кластеров 1с на сервере. Можно будет переделать на UBound(Cluster) или Clusters. Не стал заморачиваться

' Зададим пути к рабочим файлам
Const strBasesListFileName                 = "C:\zabbix\scripts\Cluster1C\Bases_List.txt"
Const strInformationBasesCountFileName     = "C:\zabbix\scripts\Cluster1C\Information_Bases_Count.txt"
Const strUsersCountFileName             = "C:\zabbix\scripts\Cluster1C\Users_Count.txt"
Const strWorkingProcessesCountFileName    = "C:\zabbix\scripts\Cluster1C\Working_Processes_Count.txt"

' Опишем переменные
Dim intInformationBasesCount            ' Количество информационных баз в кластере 1С
Dim intUsersCount                        ' Общее количество сеансов на кластере 1С
Dim intWorkingProcessesCount            ' Количество рабочих процессов в кластере 1С
Dim strBasesList                        ' Список баз

Dim Connector
Dim AgentConnection
Dim Cluster
Dim Clusters
Dim WorkingProcesses
Dim IBB
Dim IBBl
Dim Sessions
Dim SessionsActiv
Dim MyBaseName
Dim MyBaseDescr
Dim MyBaseConnectionHost
Dim SessionsActivConnection
Dim UC
Dim BC
Dim BN
Dim CL 'для подсчета номера кластера
Dim realintUsersCount 'для реального счетчика



Sub WriteToFile (ByVal strFileName, ByVal strString)

' Опишем константы для работы с текстовыми файлами
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

Dim fso, f

    Err.Clear
    On Error Resume Next

    Set fso = CreateObject("Scripting.FileSystemObject")
        Set f = fso.OpenTextFile(strFileName, ForWriting, True)
            f.WriteLine strString
            f.Close
        Set f = Nothing
    Set fso = Nothing

End Sub


intUsersCount = 0

strBasesList =  "{"
strBasesList = strBasesList & vbCrLf & "    ""data"": ["


Set Connector = CreateObject("V83.COMConnector")
    Set AgentConnection = Connector.ConnectAgent(ServerName)
        Clusters = AgentConnection.GetClusters()
		CL = 1 'Задаёт значение первого класстера
        For Each Cluster In Clusters
            AgentConnection.Authenticate Cluster, ClusterAdmin,ClusterAdminPassword
            WorkingProcesses = AgentConnection.GetWorkingProcesses(Cluster)
            intWorkingProcessesCount =  UBound(WorkingProcesses) + 1 + intWorkingProcessesCount               ' Определили количество рабочих процессов (добавил + intWorkingProcessesCount для 2+ кластеров)
            IBB = AgentConnection.GetInfoBases(Cluster)
				intInformationBasesCount = UBound(IBB) + 1 + intInformationBasesCount                           ' Определили количество баз (добавил + intInformationBasesCount для 2+ кластеров)
                BC = UBound(IBB)
                BN = 0
                For Each IBBl In IBB
                    UC = 0
                    MyBaseDescr = IBBl.Descr
                    Sessions = AgentConnection.GetInfoBaseSessions (Cluster,IBBl) 'поидее для каждой сессии есть объект License у которого есть ShortPresentation к примеру через который можно определить что за лицензия... но я так и не смог разобраться как их выдернуть
                    For Each SessionsActiv In Sessions
                        If not ((SessionsActiv.AppID = "COMConsole") or (SessionsActiv.AppID = "BackgroundJob")) Then
                            UC = UC + 1
                            intUsersCount = intUsersCount + 1                        ' Считаем пользователей в базах
						End if
					Next
					
					MyBaseDescr = Replace(IBBl.Descr,""""," ")
                    If BN = BC and CL = CLcount Then 'добавил вторую проверку CL = CLcount (подсчитывает текущий кластер для того чтобы не ставить запятую на последнем)
                            strBasesList = strBasesList & vbCrLf & "        {""{#IBDNAME}"": """ & IBBl.Name & """, ""{#IBDDESCR}"": """ & MyBaseDescr & """, ""{#IBDUSERCOUNT}"": """ & UC & """}"
                    Else
                            strBasesList = strBasesList & vbCrLf & "        {""{#IBDNAME}"": """ & IBBl.Name & """, ""{#IBDDESCR}"": """ & MyBaseDescr & """, ""{#IBDUSERCOUNT}"": """ & UC & """},"
                    End If
                    BN = BN +1
                Next
				realintUsersCount = realintUsersCount + intUsersCount 'пытаюсь определить реальный счетчик пользователей
				CL = CL + 1 'подсчитывает текущий кластер для того чтобы не ставить запятую
        Next
    Set AgentConnection = nothing
Set Connector = Nothing

strBasesList = strBasesList & vbCrLf & "    ]"
strBasesList = strBasesList & vbCrLf & "}"

' Выводим на экран результаты работы
' WScript.Echo "intUsersCount =" & intUsersCount
' WScript.Echo "intInformationBasesCount =" & intInformationBasesCount
' WScript.Echo "intWorkingProcessesCount =" & intWorkingProcessesCount
' c strBasesList


' Записываем по файлам результаты работы
WriteToFile strBasesListFileName, strBasesList
WriteToFile strInformationBasesCountFileName, intInformationBasesCount
'WriteToFile strUsersCountFileName, intUsersCount 'так было
WriteToFile strUsersCountFileName, realintUsersCount 'так стало
WriteToFile strWorkingProcessesCountFileName, intWorkingProcessesCount