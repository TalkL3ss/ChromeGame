[string]$site = "https://smfrica/main"
$chromeProfile = $env:temp+"\chrome"
$skey = $null
$url =""
$global:DayOfFill = ""

$GetCookies = @'
{
    "id": 1,
     "method": "Network.getAllCookies"
}
'@
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

function StartProc() {
    Start-Process  -FilePath "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" -ArgumentList $site,"--disable-gpu","--remote-debugging-port=9222","--user-data-dir=$chromeProfile"
    for ($i = 0; $i -le 100; $i++) {
        if (-not $(Test-NetConnection 127.0.0.1 -Port 9222).TcpTestSucceeded) { Start-Sleep -Milliseconds 100} else { break}
        }
        Hideit
    if ($(get-date).DayOfWeek -eq "Sunday") { 
        $global:DayOfFill = $((Get-Date).AddDays(-3).ToString("yyyy-MM-dd"))
    } else {
        $global:DayOfFill = $((Get-Date).AddDays(-1).ToString("yyyy-MM-dd"))
    }

    $webSocketRequest = Invoke-WebRequest -Uri http://localhost:9222/json | ConvertFrom-Json
    Write-Host 'Connect to: ' ($webSocketRequest.webSocketDebuggerUrl)[0]
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

$url = StartProc
ConnectToWebsocket -Command $LogOutParams | Out-Null
ConnectToWebsocket -Command $ConsoleParams | Out-Null
$resp = (ConnectToWebsocket -Command $GetCookies)
$resp = ($resp | ConvertFrom-Json).result.cookies
Get-CimInstance Win32_Process -Filter "CommandLine LIKE '%chrome.exe%--remote-debugging-port=9222%'" | %{Stop-Process $_.ProcessId -ErrorAction SilentlyContinue}

for ($i=0; $i -le $resp.Length; $i++ ) {
    if ($resp[$i].name -eq "skey") { $skey = $resp[$i].value }
}


if ($skey.Length -eq 0) { echo "Error! Cannot get skey!"; exit}

$session = [Microsoft.PowerShell.Commands.WebRequestSession]::new()
$cookie = [System.Net.Cookie]::new('skey', $skey)
$session.Cookies.Add('https://smfrica/', $cookie)

$baseURL = 'https://smfrica/rest'
$uri = [string]::Format('{0}{1}',$baseURL,'/users/current')
$myUserID = (Invoke-RestMethod $uri -WebSession $session).id

$uri = [string]::Format('{0}/timesheets?start_at={1}&end_at={1}&resource_id={2}',$baseURL,$DayOfFill,$myUserID)
$timeSheetID =  (Invoke-RestMethod $uri -WebSession $session).timesheet_id

$uri = [string]::Format('{0}/timesheets/{1}/assignments',$baseURL,$timeSheetID)
$headers = @{
            "Accept-Charset"= "UTF-8";
        };
$assignmentsIDs = (Invoke-RestMethod $uri -WebSession $session -Headers $headers)

$myAssignmentsIDs = @()
for ($i=0; $i -le $assignmentsIDs.Length; $i++){
    if ($assignmentsIDs[$i].closed -eq $false) {$myAssignmentsIDs += $assignmentsIDs[$i]}
    if ($assignmentsIDs[$i].days -match $DayOfFill) {echo "Day is full!"; exit}
}


$randAssignments = $myAssignmentsIDs[$(Get-Random -Minimum 0 -Maximum $myAssignmentsIDs.Length)].assignment_id
$uri = [string]::Format('{0}/timesheets/{1}/assignments/{2}/days/{3}',$baseURL,$timeSheetID,$randAssignments,$DayOfFill)


$headers = @{'accept'='application/vnd.sciforma.v1+json';'Content-Type'='application/vnd.sciforma.v1+merge-patch+json'}

Invoke-WebRequest -UseBasicParsing -Uri "https://SynerionWeb/SynerionWeb/api/personalization" -Method "POST" -Headers @{"Accept"="application/json, text/plain, */*"} -ContentType "application/json;charset=UTF-8" `
-Body "{`"DBSummaryTimeAsDecimalDisplay`":`"true`"}" -UseDefaultCredentials | Out-Null


$workedYesterday = (Invoke-restMethod -UseBasicParsing -Uri "https://SynerionWeb/SynerionWeb/api/DailyBrowser/summary" -Method "POST" -Headers @{"Accept"="application/json, text/plain, */*"} -ContentType "application/json;charset=UTF-8" `
-Body "{`"Date`":`"$($DayOfFill)T00:00:00.000Z`",`"EmployeeId`":`"       12345`",`"SummaryType`":0,`"PeriodKey`":202403,`"AccumCode`":5}" -UseDefaultCredentials).Summaries[0].Description


Invoke-RestMethod -method Patch -Uri $uri -WebSession $session -H @{'accept'='application/vnd.sciforma.v1+json';'Content-Type'='application/vnd.sciforma.v1+merge-patch+json'} `
  -Body @"
 {
    `"status`": `"WORKING`",
    `"actual_effort`": $($workedYesterday)
  }
"@


Remove-Variable DayOfFill
Remove-Variable skey
