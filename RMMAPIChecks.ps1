<#

List all Workstations
https://wwwgermany1.systemmonitor.eu.com/api/?apikey=xxx&service=list_workstations&siteid=xxx




#>
$apikey=''


[xml]$AllWorkstations=Invoke-RestMethod -Uri "https://wwwgermany1.systemmonitor.eu.com/api/?apikey=$apikey&service=list_workstations&siteid=200978" -UseBasicParsing

[xml]$AllServers=Invoke-RestMethod -Uri "https://wwwgermany1.systemmonitor.eu.com/api/?apikey=$apikey&service=list_servers&siteid=200978" -UseBasicParsing


$resultWS=foreach ($ws in $AllWorkstations.result.items.workstation) {

    $cols = $ws.ChildNodes.localname

    $props = New-Object PSObject

       For ($i=0;$i -lt $cols.count;$i++) {

           $props | Add-Member $cols[$i] $(If($ws.($cols[$i]).innertext){$($ws.($cols[$i]).innertext)}Else{$ws.($cols[$i])})

       }

    $props|Add-Member TelephoneNumber $(if($ws.user){$(try `
                                                    {(get-aduser $props.user.split("\")[-1] -Properties telephoneNumber -ErrorAction SilentlyContinue).telephoneNumber`
                                                    }`
                                                    catch{})}`
                                                    else{})

    $props|Add-Member OSBuild $((Get-ADComputer $props.name -Properties OperatingSystemVersion).OperatingSystemVersion)
    
    $props

}

$resultServers=foreach ($serv in $AllServers.result.items.server) {

    $cols = $serv.ChildNodes.localname

    $props = New-Object PSObject

       For ($i=0;$i -lt $cols.count;$i++) {

           $props | Add-Member $cols[$i] $(If($serv.($cols[$i]).innertext){$($serv.($cols[$i]).innertext)}Else{$serv.($cols[$i])})

       }

    $props|Add-Member OSBuild $((Get-ADComputer $props.name -Properties OperatingSystemVersion).OperatingSystemVersion)

    $props

}


$view_resultWS=$resultWS|select ip,name,user,os,online,TelephoneNumber,OSBuild

$view_resultServers=$resultServers|select ip,name,user,os,online,OSBuild

$resultVersion=@()

Foreach ($wsID in $resultWS.workstationid) {

$checks=Invoke-RestMethod -Uri "https://wwwgermany1.systemmonitor.eu.com/api/?apikey=$apikey&service=list_checks&deviceid=$wsID" -UseBasicParsing

$checkIDs=$null

$checkIDs=$checks.result.items.check.description|where {$_."#cdata-section" -like "*Faktura*"}|foreach {$_.parentnode.checkid}

$wsOnline=$resultWS|where {$_.workstationid -eq "$wsID"}|foreach {$_.online}

    if ($checkIDs) {

    $checkDate=$checks.result.items.check.description|where {$_."#cdata-section" -like "*Faktura*"}|foreach {$_.parentnode.date + " " + $_.parentnode.time}|Get-Date -ErrorAction SilentlyContinue

        foreach ($chID in $checkIDs) {
        
            $checkResults=Invoke-RestMethod -Uri "https://wwwgermany1.systemmonitor.eu.com/api/?apikey=$apikey&service=get_formatted_check_output&checkid=$chID" -UseBasicParsing

            $FakturaVersion=$checkResults.result.formatted_output
    
        }

    }

    else {
        
        $checkDate=""

        $FakturaVersion="No check added"
        
    }
    
$Settings=[PSCustomObject]@{
        ComputerName = ($resultWS|where {$_.workstationid -eq "$wsID"}).name
        User = ($resultWS|where {$_.workstationid -eq "$wsID"}).user
        Version = $FakturaVersion
        LastCheck = $checkDate
        Online = $wsOnline
        TelephoneNumber = ($resultWS|where {$_.workstationid -eq "$wsID"}).TelephoneNumber
    }


$resultVersion+=$Settings

}

$resultSoftware=@()

Foreach ($wsID in $resultWS.workstationid) {

    $assetID=(Invoke-RestMethod -Uri "https://wwwgermany1.systemmonitor.eu.com/api/?apikey=$apikey&service=list_device_asset_details&deviceid=$wsID" -UseBasicParsing).result.assetid

    $softwareXML=Invoke-RestMethod -Uri "https://wwwgermany1.systemmonitor.eu.com/api/?apikey=$apikey&service=list_all_software&assetid=$assetID"

    $software=foreach ($s in $softwareXML.result.items.software) {

        $cols = $s.ChildNodes.localname

        $props = New-Object PSObject

            For ($i=0;$i -lt $cols.count;$i++) {

                $props | Add-Member $cols[$i] $(
                                                If($s.($cols[$i]).innertext){
                                                
                                                $($s.($cols[$i]).innertext)
                                                }
                                                
                                                Else{$s.($cols[$i])})

            }

            $props|Add-Member "ComputerName" $(($resultWS|where {$_.workstationid -eq "$wsID"}).name)

        $props

    }

    $resultSoftware+=$software

}


$allAgents=$resultWS+$resultServers


$ADComputers=(Get-ADComputer -filter *).name



$c=Compare-Object -ReferenceObject $allagents.name -DifferenceObject $ADComputers -PassThru



$c|Get-ADComputer -Properties IPv4Address,lastBadPasswordAttempt,LastLogonDate,memberof,operatingSystem,OperatingSystemVersion|select name,IPv4Address,LastLogonDate,operatingSystem,OperatingSystemVersion,enabled,memberof,lastBadPasswordAttempt|Out-GridView