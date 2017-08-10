function ftest{
#Set-ExecutionPolicy Unrestricted -Scope CurrentUser -Force;
Import-Module "$env:ProgramFiles\Screenshot Utility\AddImagesToWordDocument.psm1";
#Get-ChildItem "C:\Users\$env:UserName\Desktop\BioData.docx" -Filter "txt.*.test.*" | Copy-Item -Destination "C:\Users\$env:UserName\Desktop\mjaintest\" -Force;
$a=$((Get-Date).DateTime);
$b=($a -replace '[,]','' -replace '[ ]','_' -replace '[:]','')+".docx";
cd "C:\Users\$env:UserName\captor_utility\";
New-Item -Name $b -ItemType File;
Add-OSCPicture -WordDocumentPath "C:\Users\$env:UserName\captor_utility\$b" -ImageFolderPath "C:\Users\$env:UserName\captor_utility\";
Remove-Item -path "C:\Users\$env:UserName\captor_utility\*" -Filter *jpg;
$filenm=Get-Content "C:\Users\Public\time.txt";
Rename-Item -path "C:\Users\$env:UserName\captor_utility\$b" -NewName "C:\Users\$env:UserName\captor_utility\$filenm";
Remove-Item -path "C:\Users\Public\time.txt";
}
&ftest