Set-DbatoolsInsecureConnection -SessionOnly
$srv = '192.168.50.40'
if (-not $cred) {
    $cred = Get-Credential -UserName sa -Message '..'
}

$result = Get-DbaDatabase -SqlInstance $srv -SqlCredential $cred -ExcludeSystem | Out-GridView -PassThru | Invoke-DbaQuery -File I:\InactiveVBClients\get_last_UpdBnd.sql

$result
