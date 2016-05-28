<#
.SYNOPSIS
PermissionMatrixGraphBuilder.ps1 - easy script that generates graph from provided CSV file.

.DESCRIPTION 
Easy script that generates graph from provided CSV file.

.OUTPUTS
Output set in the script is a png file, however that can be changed by modifying -T parameter for dot.exe

.PARAMETER Users
This parameter - when set to $true - accepts couple of users at one time, put there user you would like to find dependencies for.

.PARAMETER Level
It determines how many times sript goes through the CSV file given as the input.
Giving level as 2 means that script will go through the file one time, and later second time 
searching for depencies for next batch of users.

.PARAMETER SingleUser
That parameter means that it will only read dependencies for one user.


.EXAMPLE
.\PermissionMatrixGraphBuilder.ps1 -users "Pawel Jarosz", "Test User", "Next User"
That will create a graph with dependencies for 3 users given in "Users" parameter

.EXAMPLE
.\PermissionMatrixGraphBuilder.ps1 -users "Pawel Jarosz", "Test User", "Next User" -Level 1
That will do same as above :) with only difference - it will go though the file only once, so
even if for example "Teddy Burns" will have access to "Pawel Jarosz" mailbox
, and there will be some dependencies to that mailbox but entries will be BEFORE entry with
"Teddy Burns" -> "Pawel Jarosz" permissions - it won't be shown, howeven - level 2 might show it

.EXAMPLE
.\PermissionMatrixGraphBuilder.ps1 -users "Pawel Jarosz", "Test User", "Next User" -SingleUser $true
That will show only mailboxes that have direct connection to mentioned in "Users" array

.LINK
https://paweljarosz.wordpress.com/2016/05/28/exchange-mailboxfolders-permissions-dependency-graph-between-users

.NOTES
Written By: Paweł Jarosz

Find me on:
* My Blog:	https://paweljarosz.wordpress.com/
* LinkedIn:	https://www.linkedin.com/in/paweljarosz2
* GoldenLine: 	http://www.goldenline.pl/pawel-jarosz2/
* Github:	https://github.com/zaicnupagadi


Change Log:
V1.00, 01/05/2016 - Initial version

#>

param(
	[Parameter(ParameterSetName='mailbox')] [string[]]$users,
    [int] $Level,
    [string] $SingleUser
)

$GraphImageFile = "GraphImageFile.png"
$GraphGraphVizFile = "GraphVizFile.gv"
$CSVPermissionsFile = "Permissions.csv"

$global:report = @()
$global:DepArray = @()
$global:a = import-csv $CSVPermissionsFile -Delimiter ";"

if (!$Users){
    ForEach ($p in $global:a){
        if ($Users -notcontains $p.mailbox ){ 
        $users += $p.mailbox
        }
    }
}

$users

if ($level) {$global:i=0} else {$global:i=666}

Function CheckPermissions ([string[]]$mailbox) {
$IsChanged = $global:DepArray
    ForEach ($mail in $mailbox){
        ForEach ($b in $global:a) {
            if (($b.mailbox -eq $mail -and $b.mailbox -ne "") -or ($b.user -eq $mail -and $b.user -ne "")) {
                if ($global:DepArray -notcontains $b.mailbox){
                $global:DepArray += $b.mailbox
                }
                if ($global:DepArray -notcontains $b.user){
                $global:DepArray += $b.user
                }
    
            }
        }

    }
    $global:i++
    if ((compare $IsChanged $global:DepArray) -and ($global:i -ne $level)) {
       CheckPermissions ([string[]]$global:DepArray)    
    } else {
        ForEach ($x in $global:a){
            ForEach ($l in $DepArray){
                if (!$SingleUser) {
                    if ($x.mailbox -eq $l -or $x.User -eq $l){
                    $global:report += $x
                    }
                    } else {
                    if ($x.mailbox -eq $l -or $x.User -eq $l -and ($Users -contains $x.mailbox -or $Users -contains $x.User)){
                    $global:report += $x
                    }
                    
                }
            }
        }
    }

}

CheckPermissions ([string[]]$users)
$PermReport = $global:report | sort mailbox,User -Unique

"digraph{" > $GraphGraphVizFile
ForEach ($Entry in $PermReport) {
    if ($Entry.Mailbox -and $Entry.User){
        if ($users -contains $Entry.Mailbox){
        $L = '"'+$Entry.Mailbox+'"'+'[fillcolor="#118EC4",style="filled"]'
        $L >> $GraphGraphVizFile
        $L = '"'+$Entry.User+'"'+"->"+'"'+$Entry.mailbox+'"'+';'
        $L >> $GraphGraphVizFile
        } elseif ($users -contains $Entry.user) {
        $L = '"'+$Entry.User+'"'+'[fillcolor="#118EC4",style="filled"]'
        $L >> $GraphGraphVizFile
        $L = '"'+$Entry.User+'"'+"->"+'"'+$Entry.mailbox+'"'+';'
        $L >> $GraphGraphVizFile
        } else {
        #$L = $Entry.Mailbox+"->"+$Entry.User+' [label="PATH"];'
        $L = '"'+$Entry.User+'"'+"->"+'"'+$Entry.mailbox+'"'+';'
        $L >> $GraphGraphVizFile
        }
    }
}
"}" >> $GraphGraphVizFile
get-content $GraphGraphVizFile | & 'C:\Program Files (x86)\Graphviz2.38\bin\dot.exe' -Tpng -o "$GraphImageFile"
ii "$GraphImageFile"