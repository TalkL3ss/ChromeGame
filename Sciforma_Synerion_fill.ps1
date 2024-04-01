[string]$site = "https://Sciforma/Sciformaprod/"
$baseURL = [string]::Format('{0}{1}{2}',"https://",$($site.Split('/')[2]),"/Sciformaprod/rest")
$baseCookie = [string]::Format('{0}{1}{2}',"https://",$($site.Split('/')[2]),"/")
$snryion = 'https://Synerion/SynerionWeb'
$chromeProfile = $env:temp+"\chrome"
$global:DayOfFill = ""
$global:skx=0

#Documention From
#https://chromedevtools.github.io/devtools-protocol/tot/Network/

$GetCookies = @'
{
    "id": 1,
     "method": "Network.getAllCookies"
}
'@
#search the login button and do the login
$ConsoleParams = @'
{
    "id": 1,
     "method": "Runtime.evaluate",
     "params": {
            "expression": "setInterval(() => {if (document.body.innerHTML.includes('login/buttonSSO')) { document.getElementById('login/buttonSSO').click();  }}, 1000);"
        }
}
'@
$LogOutParams = @'
{
    "id": 1,
     "method": "Runtime.evaluate",
     "params": {
            "expression": "setInterval(() => {if (document.body.innerHTML.includes('usermenu/logout')) { document.getElementById('usermenu/logout').click()  }}, 1000);"
        }
}
'@

Function Hideit() {
#Hide the Chrome
$definition = @"
    [DllImport("user32.dll")]
    static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
    
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    public static void Show(string wClass, string wName)
    {
        IntPtr hwnd = FindWindow(wClass, wName);
        if ((int)hwnd > 0)
            ShowWindow(hwnd, 1);
    }

    public static void Hide(string wClass, string wName)
    {
        IntPtr hwnd = FindWindow(wClass, wName);
        if ((int)hwnd > 0)
            ShowWindow(hwnd, 0);
    }
"@
Add-Type -MemberDefinition $definition -Namespace my -Name WinApi
[my.WinApi]::hide('Chrome_WidgetWin_1', 'Untitled - Google Chrome')
[my.WinApi]::hide('Chrome_WidgetWin_1', 'Restore pages?')
}

Function StartProc() {
#Start Chrome with devtool
    for ($i = 0; $i -le 100; $i++) {
        if (-not $(Test-NetConnection $($site.Split('/')[2]) -Port 443).TcpTestSucceeded) { echo "Wait for $($site.Split('/')[2]) $i time";Start-Sleep -Seconds 60} else {break}
        if ($i -eq 100) {Write-Host "Looks like sciforma is down"; exit}
     }
    Start-Process  -FilePath "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" -ArgumentList $site,"--disable-gpu","--remote-debugging-port=9222","--user-data-dir=$chromeProfile"
    for ($i = 0; $i -le 100; $i++) {
        if (-not $(Test-NetConnection 127.0.0.1 -Port 9222).TcpTestSucceeded) { Start-Sleep -Seconds 2} else { break}
        }
        if (!$psISE) {Hideit}
    #Create a varible when to update
    if ($(get-date).DayOfWeek -eq "Sunday") { 
        $global:DayOfFill = $((Get-Date).AddDays(-3).ToString("yyyy-MM-dd"))
        $global:MonthOfFill = $((Get-Date).AddDays(-3).ToString("yyyyMM"))
    } else {
        $global:DayOfFill = $((Get-Date).AddDays(-1).ToString("yyyy-MM-dd"))
        $global:MonthOfFill = $((Get-Date).AddDays(-1).ToString("yyyyMM"))
    }

    $webSocketRequest = Invoke-WebRequest -Uri http://localhost:9222/json | ConvertFrom-Json
    Write-Host 'Connect to: ' ($webSocketRequest.webSocketDebuggerUrl)[0]
    #return the url for control the devtool
    return ($webSocketRequest.webSocketDebuggerUrl)[0]
}

Function ConnectToWebsocket($Command){
    $ws = New-Object System.Net.WebSockets.ClientWebSocket
    $cancelationtonken = New-Object System.Threading.CancellationToken
    $Buffer_Size = 2048

    $ConsoleParamsL = @()
    $Command.ToCharArray() | % {$ConsoleParamsL += [byte] $_}          
    $ConsoleParamsP = New-Object System.ArraySegment[byte]  -ArgumentList @(,$ConsoleParamsL)
    $ConsoleParamsL = [byte[]] @(,0) * $Buffer_Size
    $ConsoleParamsR = New-Object System.ArraySegment[byte]  -ArgumentList @(,$ConsoleParamsL)

    $Connect = $ws.ConnectAsync($url, $cancelationtonken)

    While (!$Connect.IsCompleted) { Start-Sleep -Milliseconds 100 }
    $Connect = $ws.SendAsync($ConsoleParamsP, [System.Net.WebSockets.WebSocketMessageType]::Text, [System.Boolean]::TrueString, $cancelationtonken)
    while (!$Connect.IsCompleted) { Start-Sleep -Milliseconds 100 }
    sleep -Seconds 5
    if ($ws.State -eq 'Open') {
        do {
            $Connect = $ws.ReceiveAsync($ConsoleParamsR, $cancelationtonken)
                While (!$Connect.IsCompleted) { Start-Sleep -Milliseconds 100 }
            $ConsoleParamsR.Array[0..($Connect.Result.Count - 1)] | ForEach { $ret += [char]$_ }
        } until ($Connect.Result.Count -lt $Buffer_Size)
        }
        return $ret
    $Connect = $ws.CloseAsync([System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure, "NormalClosure", $cancelationtonken)

}

Function GetSKey() {
ConnectToWebsocket -Command $LogOutParams | Out-Null #send logout to get fresh token
ConnectToWebsocket -Command $ConsoleParams | Out-Null #send login to get fresh token
#get the skey authoreztion for the maof api
    $resp = (ConnectToWebsocket -Command $GetCookies)
    $resp = ($resp | ConvertFrom-Json).result.cookies
    

    for ($i=0; $i -le $resp.Length; $i++ ) {
        if ($resp[$i].name -eq "skey") { $skey = $resp[$i].value }
    }

   
    if (($skey.Length -eq 0) -and ($global:skx -ge 10)){
        Get-CimInstance Win32_Process -Filter "CommandLine LIKE '%chrome.exe%--remote-debugging-port=9222%'" | %{Stop-Process $_.ProcessId -ErrorAction SilentlyContinue} 
        Write-host "Error! Cannot get skey!"
        exit
     } elseif ($skey.Length -eq 0)  {$global:skx=$global:skx+1; Write-Host "Cannot get skey, another try in 1 sec, try #:"$($global:skx); Start-Sleep -Seconds 1 ;GetSKey}
     else { return $skey }
}

function SynerionStuff() {
    Invoke-WebRequest -UseBasicParsing -Uri "$($snryion)/api/personalization" -Method "POST" -Headers @{"Accept"="application/json, text/plain, */*"} -ContentType "application/json;charset=UTF-8" `
    -Body "{`"DBSummaryTimeAsDecimalDisplay`":`"true`"}" -UseDefaultCredentials | Out-Null #set the Synerion to output Decimal to get the the proper date format

    $EmployID = (Invoke-WebRequest -UseBasicParsing -Uri "$($snryion)/api/DailyBrowser/Attendance" -Method "POST" -Headers @{"Accept"="application/json, text/plain, */*"} `
    -ContentType "application/json;charset=UTF-8" -Body "{`"Employees`":null,`"SelectionMode`":0,`"DatePeriodSelection`":{`"AccumCode`":5,`"DateRange`":{`"From`":`"$($DayOfFill)T00:00:00.000Z`",`"To`":`"$($DayOfFill)T00:00:55.000Z`"},`"IsDateRange`":false,`"PeriodKey`":$($MonthOfFill)},`"FirstResult`":0,`"ItemsOnPage`":1,`"SortDescriptors`":null,`"Filters`":null,`"LoadEmployeeMode`":1}" -UseDefaultCredentials)

    $EmployID = ($EmployID.Content | ConvertFrom-Json).DailyBrowserDtos[0].EmployeeId #get employeeid


    $workedYesterday = (Invoke-restMethod -UseBasicParsing -Uri "$($snryion)/api/DailyBrowser/summary" -Method "POST" -Headers @{"Accept"="application/json, text/plain, */*"} -ContentType "application/json;charset=UTF-8" `
    -Body "{`"Date`":`"$($DayOfFill)T00:00:00.000Z`",`"EmployeeId`":`"$($EmployID)`",`"SummaryType`":0,`"PeriodKey`":$($MonthOfFill),`"AccumCode`":5}" -UseDefaultCredentials).Summaries[0].Description
    #check how much did i worked yesterday

    if ($workedYesterday.Length -eq 0) {
        Write-Host "Looks like you didn't worked at "$($DayOfFill)" or the day didn't close in "$($snryion)
        Get-CimInstance Win32_Process -Filter "CommandLine LIKE '%chrome.exe%--remote-debugging-port=9222%'" | %{Stop-Process $_.ProcessId -ErrorAction SilentlyContinue}
        exit
    }
    return $workedYesterday
}


$url = StartProc
$workedYesterday = SynerionStuff
$skey= GetSKey #get the token
Get-CimInstance Win32_Process -Filter "CommandLine LIKE '%chrome.exe%--remote-debugging-port=9222%'" | %{Stop-Process $_.ProcessId -ErrorAction SilentlyContinue}

$session = [Microsoft.PowerShell.Commands.WebRequestSession]::new()
$cookie = [System.Net.Cookie]::new('skey', $skey)
$session.Cookies.Add($baseCookie, $cookie)


$uri = [string]::Format('{0}{1}',$baseURL,'/users/current')
$myUserID = (Invoke-RestMethod $uri -WebSession $session).id #get userid in maof

$uri = [string]::Format('{0}/timesheets?start_at={1}&end_at={1}&resource_id={2}',$baseURL,$DayOfFill,$myUserID)
$timeSheetID =  (Invoke-RestMethod $uri -WebSession $session).timesheet_id #get timesheet

$uri = [string]::Format('{0}/timesheets/{1}/assignments',$baseURL,$timeSheetID)
$assignmentsIDs = (Invoke-RestMethod $uri -WebSession $session) #get all the assignments in the maof

$myAssignmentsIDs = @()
for ($i=0; $i -le $assignmentsIDs.Length; $i++){
    if ($assignmentsIDs[$i].closed -eq $false) {$myAssignmentsIDs += $assignmentsIDs[$i]} #get only open close assgiments 
    if ($assignmentsIDs[$i].days -match $DayOfFill) {write-host "Day is full!"; exit} #check if the needed day allready full
}


$randAssignments = $myAssignmentsIDs[$(Get-Random -Minimum 0 -Maximum $myAssignmentsIDs.Length)].assignment_id
$uri = [string]::Format('{0}/timesheets/{1}/assignments/{2}/days/{3}',$baseURL,$timeSheetID,$randAssignments,$DayOfFill)


Invoke-RestMethod -method Patch -Uri $uri -WebSession $session -H @{"accept"="application/vnd.sciforma.v1+json";"Content-Type"="application/vnd.sciforma.v1+merge-patch+json"}  `
  -Body @"
 {
    `"status`": `"WORKING`",
    `"actual_effort`": $($workedYesterday)
  }
"@  #do the update in sciforma

if ($psISE) {
    Remove-Variable DayOfFill
    Remove-Variable skey
}
