# Importar o módulo ImportExcel
Import-Module -Name ImportExcel

# Caminho para o arquivo Excel
$ExcelFilePath = "C:\Users\admin.gonzales\Desktop\usuarios_ativos.xlsx"

# Nome do grupo no AD
$GroupName = "Usuários GED"

# Carregar a planilha Excel
$Worksheet = Import-Excel -Path $ExcelFilePath

# Percorrer as linhas da coluna "Usuario" e adicionar ao grupo
foreach ($Row in $Worksheet) {
    $Username = $Row.Usuario

    if ($Username) {
        $User = Get-ADUser -Filter {SamAccountName -eq $Username}
        
        if ($User) {
            $Group = Get-ADGroup $GroupName
            if ($Group) {
                Add-ADGroupMember -Identity $Group -Members $User
                Write-Host "O usuário $Username foi adicionado ao grupo $GroupName."
            } else {
                Write-Host "O grupo $GroupName não foi encontrado no Active Directory."
            }
        } else {
            Write-Host "O usuário $Username não foi encontrado no Active Directory."
        }
    }
}
