$msi = New-Object -ComObject WindowsInstaller.Installer
$db = $msi.OpenDatabase("path\to\msi", 0)
$view = $db.OpenView("SELECT Value FROM Property WHERE Property = 'ProductCode'")
$view.Execute()
$record = $view.Fetch()
$record.StringData(1)
