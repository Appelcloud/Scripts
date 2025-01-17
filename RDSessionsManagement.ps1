param (
    [Parameter(Mandatory=$true, ParameterSetName="Enable")]
    [switch]$EnableNewSessions,

    [Parameter(Mandatory=$true, ParameterSetName="Disable")]
    [switch]$DisableNewSessions,

    [Parameter(Mandatory=$true, ParameterSetName="Disable", DontShow, HelpMessage="Type either 'All' or 'Disconnected'")]
    [ValidateSet("All", "Disconnected")]
    [string]$Disconnect
)

$currentdate = Get-Date -Format "mm-HH-ss_dd-MM-yyyy"
Start-Transcript -Path "C:\Windows\Logs\ScriptLogs\RDSSessionHosts-$currentdate.txt"

if (-not (Get-Module -Name RemoteDesktop)) {
    Write-Host "Importing RemoteDesktop module..."
    Import-Module RemoteDesktop
} else {
    Write-Host "RemoteDesktop module is already imported."
}

# Read the server list from the file
$servers = Get-Content -Path "####.txt"

# Loop through each server in the list
foreach ($server in $servers) {
    Write-Host "Processing server: $server"

    if ($EnableNewSessions) {
        Write-Host "Enabling new sessions on server: $server"
        Set-RDSessionHost -SessionHost $server -NewConnectionAllowed Yes
    }

    if ($DisableNewSessions) {
            Write-Host "Disabling new sessions on server: $server"
            Set-RDSessionHost -SessionHost $server -NewConnectionAllowed No
            # Get the list of sessions

        $sessions = query session /server:$server | ForEach-Object {
            $fields = $_ -split ' +'
            [PSCustomObject]@{
                SessionId = $fields
            }
        }

        $sessionsDisconnected = query session /server:$server | Where-Object { $_ -match "Disc" } | ForEach-Object {
            $fields = $_ -split ' +'
            [PSCustomObject]@{
                SessionId = $fields
            }
        }

        # Determine which sessions to log off based on the disconnect parameter
        if ($PSCmdlet.ParameterSetName -eq "Disable") {
            if ($Disconnect -eq "All") {
                $sessionsToLogOff = $sessions
            } else {
                $sessionsToLogOff = $sessionsDisconnected
            }
        } else {
            $sessionsToLogOff = $sessions
        }

        # Log off each selected session
        foreach ($session in $sessionsToLogOff) {
            Write-Host "Logging off user: $($session.SessionId[1]) with session id: $($session.SessionId[2]) on server: $server"
            logoff $session.SessionId[2] /server:$server
        }
    }
}
Stop-Transcript
