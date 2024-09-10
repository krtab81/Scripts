Get-ChildItem *.msg -Exclude "????-??-?? *" |
ForEach-Object{
    $outlook = New-Object -comobject outlook.application
    $msg = $outlook.CreateItemFromTemplate($_.FullName)
    $msg | Select senderemailaddress,to,subject,Senton,body|ft -AutoSize
    Rename-Item -LiteralPath $_.FullName -NewName "$($msg.Senton.ToString('yyyy-MM-dd')) $($_.Basename)_$($_.Extension)"
    }