Throw "Safety"

Add-Type -assembly "Microsoft.Office.Interop.Outlook"
$Outlook = New-Object -comobject Outlook.Application
$namespace = $Outlook.GetNameSpace("MAPI")

$inbox = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)
$emailAddress = Read-Host "What is your email address"
$PrimaryFolders = $namespace.Folders.Item($emailAddress).Folders




#region #### EMPTY ALL SUBFOLDERS IN A SPECIFIC FOLDER INTO THE ARCHIVE FOLDER ####
$OrigArchive = $namespace.Folders.Item().Folders.ite('Archive').folders.item('(4) FYI').folders
$foldersToMove = $OrigArchive | Select-Object name | Sort-Object name | Select-Object -First 25
$NewArchive = $namespace.Folders.Item().Folders.ite('Archive')

foreach($folder in $foldersToMove){
    $name = $folder.Name
    $archiveFolder = $namespace.Folders.Item().Folders.ite('Archive').folders.item('(4) FYI').folders.item($name)
    $archiveFolder.items.count
    }
    $archiveFolder.findnext
#endregion


#region #### COUNT HOW MANY EMAILS FROM EACH PERSON IN INBOX ####
$allSenders = $inbox.Items | Select-Object -Property sendername -Unique
foreach($sender in $allSenders){
    $items = $inbox.Items | Where-Object {$_.sendername -eq $sender}
    $items
    }
#endregion