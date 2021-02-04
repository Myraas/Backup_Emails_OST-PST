# It's sloppy code, but it gets the job done! I plan on improving this later and appending .old to existing files in the backup location,
# that way this can be run from task scheduler on a regular basis and provide multiple versions. This will make email migrations a breeze!

Clear-Host

$destination = "\\FileServer\BackupLocation"


$explorerprocesses = @(Get-WmiObject -Query "Select * FROM Win32_Process WHERE Name='explorer.exe'" -ErrorAction SilentlyContinue)
ForEach ($i in $explorerprocesses)
    {
        $Username = $i.GetOwner().User
        taskkill /F /T /IM outlook.exe /FI “USERNAME eq $Username”
        }

$Computername = hostname
pushd $destination
mkdir $computername
popd

$Userprofiles = @(Get-ChildItem -Path C:\Users -Name)
Foreach ($user in $Userprofiles)
    {
        pushd $destination\$computername
        mkdir $user
        popd
        Copy-Item -Path "C:\Users\$user\AppData\Local\Microsoft\Outlook\*.ost" -Destination "$destination\$computername\$user" -Recurse
        Copy-Item -Path "C:\Users\$user\AppData\Local\Microsoft\Outlook\*.pst" -Destination "$destination\$computername\$user" -Recurse
        }
