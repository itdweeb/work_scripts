function get-dn ($SAMName)    {
    $root = [ADSI]''
    $searcher = new-object System.DirectoryServices.DirectorySearcher($root)
    $searcher.filter = "(&(objectClass=user)(sAMAccountName= $SAMName))"
    $user = $searcher.findall()

    if ($user.count -gt 1)      {     
            $count = 0
                foreach($i in $user)            { 
            write-host $count ": " $i.path 
                    $count = $count + 1
                }

            $selection = Read-Host "Please select item: "
        return $user[$selection].path

          }      else      { 
          return $user[0].path
          }
}

$Name = $args[0]
$status = $args[1]
$path = get-dn $Name

if ($path -ne $null)    {

    "'" + $path + "'"  
    if ($status -match "enable")     {
        # Enable the account
        $account=[ADSI]$path
        $account.psbase.invokeset("AccountDisabled", "False")
        $account.setinfo()
    }    else    {
        # Disable the account
        $account=[ADSI]$path
        $account.psbase.invokeset("AccountDisabled", "True")
        $account.setinfo()
    }
}    else    {
    write-host "No user account found!" -foregroundcolor white -backgroundcolor red
}