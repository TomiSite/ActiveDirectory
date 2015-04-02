############################################
#Author:Lixiaosong
#Email:lixs@ourgame.com;lixiaosong8706@gmail.com
#For:根据HAB自动调整OU结构
#Version:1.0
##############################################
Import-Module activedirectory
$AdGroups=Get-ADGroup -SearchBase "OU=Addresslist,DC=GLOA,DC=com" -Properties * -filter * |%{$_.name}
# 1 根据HAB通讯组创建所有新的部门OU
foreach ($group in $adgroups){
if ((Get-ADOrganizationalUnit -Filter "Name -like '$group'") -eq $null){New-ADOrganizationalUnit -name $Group -path "OU=OurGame,DC=GLOA,DC=com" -ProtectedFromAccidentalDeletion $false
write-host "部门OU $group 已创建" -ForegroundColor Yellow
}
}
write-host "------------------------所有部门Ou创建完毕--------------------------------" -ForegroundColor White
foreach ($AdGroup in $AdGroups){

    #2 根据HAB组隶属关系调整Ou组织结构               
    #本部门上一级部门组OU
    $DepOnOu=Get-ADGroup $AdGroup -Properties * |%{$_.memberof}
    echo "DepOnOu $DepOnOu"
    #本部门上一级部门名称
    $DepOnOuName=Get-ADGroup -SearchBase "$DepOnOu" -filter * |%{$_.Name}
    echo "DepOnOuName $DepOnOuName"
    #本部门OU
    $DepOuPath=(Get-ADOrganizationalUnit -Filter "Name -like '$AdGroup'").distinguishedname
    echo "DepOuPath $DepOuPath"
    #取消Ou的防删除保护选项以便进行迁移
    Set-ADOrganizationalUnit $DepOuPath -ProtectedFromAccidentalDeletion $false
    #上一级OU
    $DepOnOuPath=(Get-ADOrganizationalUnit -Filter "Name -like '$DepOnOuName'").distinguishedname
    echo "deponoupath $depOnOuPath"
    Set-ADOrganizationalUnit $DepOnOuPath -ProtectedFromAccidentalDeletion $false
    #迁移本部门OU到上一级部门OU
    Move-ADObject "$DepOuPath" -TargetPath "$DepOnOuPath"
    write-host "部门路径$DepOuPath 已经迁移到 $DepOnOupath " -ForegroundColor Blue
}
write-host "------------------------所有部门Ou已迁移至上级OU--------------------------------" -ForegroundColor White
foreach ($SeGroup in $AdGroups){
    #3 根据HAB组成员迁移账号至新的相应OU
    $NewDepOuPath=(Get-ADOrganizationalUnit -Filter "Name -like '$SeGroup'").distinguishedname
    #查找通讯组用户成员
    $GroupMembers=Get-ADGroupMember $SeGroup |Where-Object {$_.objectclass -like "user"}
    foreach ($GroupMember in $GroupMembers){
    Move-ADObject "$GroupMember" -TargetPath "$NewDepOuPath"
    write-host "$groupMember 已经迁移到 $NewDepOuPath " -ForegroundColor Cyan
        }

    #4 根据Exchange HAB创建安全组
    $Mailgroup=Get-ADGroup $SeGroup -Properties * |%{$_.mailNickname}
    echo "$mailgroup"
    $SecurityGroup=$mailgroup+".Se"
    echo "$securitygroup"

    #创建安全组
    New-ADGroup -Name "$SecurityGroup" -SamAccountName "$SecurityGroup"  -GroupScope Global -DisplayName "$SecurityGroup" -Path "$NewDepOuPath" -Description $SeGroup
    write-host "$SeGroup 安全组 $SecurityGroup 已经创建" -ForegroundColor Green

    #5 根据Ou结构将Ou成员加入对应的新安全组
    $NewSeGroups=Get-ADGroup -SearchBase "$NewDepOuPath" -Searchscope onelevel -Filter *
    $OuUsers=Get-ADUser -SearchBase "$NewDepOuPath" -Filter *

    foreach ($NewSeGroup in $NewSeGroups){
        foreach ($OuUser in $OuUsers){
                 Add-ADGroupMember $NewSeGroup -Members $OuUser
                 write-host "$OuUser 已经加入安全组 $NewSeGroups" -ForegroundColor Magenta
    }
    }   
}
write-host "------------------------所有部门账号已迁移至部门OU----------------------------------------" -ForegroundColor White

write-host "------------------------所有部门安全组已创建并归属到部门OU--------------------------------" -ForegroundColor White

write-host "------------------------所有部门成员已加入部门安全组--------------------------------------" -ForegroundColor White

#删除无成员OU
$OurgameOus=Get-ADOrganizationalUnit -SearchBase "OU=OurGame,DC=gloa,DC=com" -filter * |%{$_.DistinguishedName}
Foreach ($OurgameOu in $OurgameOus){
    Set-ADOrganizationalUnit $OurgameOu -ProtectedFromAccidentalDeletion $false
    $users=Get-Aduser -SearchBase "$OurgameOu" -Filter *
    if ($users -eq $null){
        Move-ADObject $OurgameOu -TargetPath "OU=OurGame_Old,DC=gloa,DC=com"
        Write-host "$OurgameOu 已经被迁移" -ForegroundColor Green
    }
          
}

write-host "------------------------所有无账号OU已迁移--------------------------------" -ForegroundColor White

#恢复OU防删除保护
Foreach ($OU in $OurgameOus){
      Set-ADOrganizationalUnit $OU -ProtectedFromAccidentalDeletion $true
      Write-host "$OU 防删除功能已开启" -ForegroundColor Green
}
write-host "------------------------所有OU防删除功能已开启--------------------------------" -ForegroundColor White
