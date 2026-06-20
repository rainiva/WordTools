$w = New-Object -ComObject Word.Application
$w.Visible = $true
Start-Sleep -Seconds 5
foreach ($a in @($w.COMAddIns)) {
    if ($a.ProgId -eq 'WordTools.ThisAddIn') {
        Write-Output "Connect=$($a.Connect)"
        Write-Output "ObjectIsNull=$($null -eq $a.Object)"
        if ($null -ne $a.Object) {
            Write-Output "Type=$($a.Object.GetType().FullName)"
            $methods = $a.Object.GetType().GetMethods() | Where-Object { $_.Name -like '*Automation*' } | Select-Object -ExpandProperty Name
            Write-Output "AutomationMethods=$methods"
        }
    }
}
$w.Quit()
