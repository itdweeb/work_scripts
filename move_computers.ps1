import-module ActiveDirectory
get-content c:\old-computers.txt | Move-ADObject -TargetPath 'cn=computers,dc=chi,dc=goldwind,dc=local'