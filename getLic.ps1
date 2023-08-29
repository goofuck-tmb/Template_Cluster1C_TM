[CmdletBinding()]
Param(
 [parameter()][alias("f")] $OutputFileName = "",
 [parameter()][alias("s")] $Server = [System.Net.Dns]::GetHostEntry([string]"localhost").HostName,
 [parameter()][alias("u")] $ClusterAdminName = "",
 [parameter()][alias("p")] $ClusterAdminPass = "",
 [parameter()][alias("o")] $OutputFormat = "csv",
 [parameter()][alias("d")] $Delimeter = ",",
 [switch] $AppendTotal = $False
)

$ServerAddress = "tcp://$Server"
$Tab = " " *4
$CountServLic = 0
$CountTSLic = 0


try {
 $V8Com = New-Object -COMObject "V83.COMConnector"
 $ServerAgent = $V8Com.ConnectAgent($ServerAddress)
} catch {
 try {
  $V8Com = New-Object -COMObject "V82.COMConnector"
  $ServerAgent = $V8Com.ConnectAgent($ServerAddress)
 } catch {
  Write-Error "Ни V82, ни V83 .COMConnector`ы не найдены, либо сервер не отвечает."
  Exit
 }
}

$TotalList = @()
ForEach ($Cluster in $ServerAgent.GetClusters()) {
 $ServerAgent.Authenticate($Cluster,$ClusterAdminName,$ClusterAdminPass)
 $CurrentCluster = New-Object PSCustomObject -Property @{
  "Tag" = "Cluster"
  "ClusterName" = $Cluster.ClusterName
  "HostName" = $Cluster.HostName
  "MainPort" = $Cluster.MainPort
  "IP" = ([System.Net.Dns]::GetHostAddresses("$Server") | Where-Object {$_.AddressFamily -eq "InterNetwork"} | Select-Object IPAddressToString -ExpandProperty IPAddressToString)
  "Bases" = @()
 }
 ForEach ($Base in $ServerAgent.GetInfoBases($Cluster)) {
  $CurrentBase = New-Object PSCustomObject -Property @{
   "Tag" = "Base"
   "BaseName" = $Base.Name
   "Sessions" = @()
  }
  ForEach ($Session in $ServerAgent.GetInfoBaseSessions($Cluster, $Base)) {
	  
	  if ($Session.License.ShortPresentation -like '*Лицензия 1*' -or $Session.License.ShortPresentation -like '*Лицензия 2*' -or $Session.License.ShortPresentation -like '*Лицензия 3*')
		{
		$CountServLic = 1 + $CountServLic
		}

		elseif ($Session.License.ShortPresentation -like "*Лицензия 4*" -or $Session.License.ShortPresentation -like "*Лицензия 5*")
		{
		$CountTSLic = 1 + $CountTSLic
		}
	  
	  
   $CurrentSession = New-Object PSCustomObject -Property @{
    "Tag" = "Session"
    "userName" = $Session.userName
    "AppID" = $Session.AppID
    "Host" = $Session.Host
    #"StartedAt" = $Session.StartedAt
    #"SessionID" = $Session.SessionID
    "Licenses" = @()
   }
   if ($Session.License) {
    try {
     $CurrentLicense = New-Object PSCustomObject -Property @{
      "Tag" = "License"
      "FileName" = $Session.License.FileName
      #"FullPresentation" = $Session.License.FullPresentation
      "IssuedByServer" = $Session.License.IssuedByServer
      "LicenseType" = $Session.License.LicenseType
      "MaxUsersAll" = $Session.License.MaxUsersAll
      "MaxUsersCur" = $Session.License.MaxUsersCur
      "Net" = $Session.License.Net
      "RMngrAddress" = $Session.License.RMngrAddress
      "RMngrPID" = $Session.License.RMngrPID
      "RMngrPort" = $Session.License.RMngrPort
      "Series" = $Session.License.Series
      "ShortPresentation" = $Session.License.ShortPresentation
	     }
	 $CurrentSession.Licenses += $CurrentLicense
	 
    } catch {}
   }
   $CurrentBase.Sessions += $CurrentSession
  }
  $CurrentCluster.Bases += $CurrentBase
 }
 $TotalList += $CurrentCluster
}
echo $CountServLic | Out-File -FilePath C:\zabbix\scripts\Cluster1C\LicensesServCount.txt
echo $CountTSLic | Out-File -FilePath C:\zabbix\scripts\Cluster1C\LicensesTSCount.txt