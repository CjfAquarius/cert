# 强制导入证书并使用弹窗显示结果
# 需要在管理员权限下运行

# 1. 指定证书文件路径
$certFilePath = "main.cer"

# 2. 创建 WScript.Shell 对象用于弹窗
$wshell = New-Object -ComObject WScript.Shell

# 3. 检查管理员权限
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    $wshell.Popup("必须使用管理员权限运行 PowerShell！`n`n请右键点击 PowerShell -> 以管理员身份运行", 0, "错误", 0 + 16)
    exit 1
}

# 4. 检查证书文件是否存在
if (-NOT (Test-Path $certFilePath)) {
    $wshell.Popup("证书文件不存在：`n$certFilePath", 0, "文件未找到", 0 + 16)
    exit 1
}

$certInfo = ""
try {
    # 5. 加载证书
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
    $cert.Import($certFilePath)
    
    $certInfo = "证书信息：`n`n"
    $certInfo += "主题：$($cert.Subject)`n"
    $certInfo += "颁发者：$($cert.Issuer)`n"
    $certInfo += "有效期：$($cert.NotBefore.ToString('yyyy-MM-dd')) 到 $($cert.NotAfter.ToString('yyyy-MM-dd'))`n"
    $certInfo += "指纹：$($cert.Thumbprint)`n"
    $certInfo += "`n状态："
    
    # 6. 打开本地计算机的根证书存储
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "LocalMachine")
    $store.Open("ReadWrite")
    
    # 7. 检查是否已存在
    $existingCerts = $store.Certificates.Find([System.Security.Cryptography.X509Certificates.X509FindType]::FindByThumbprint, $cert.Thumbprint, $false)
    
    if ($existingCerts.Count -gt 0) {
        $certInfo += "检测到已存在相同指纹的证书`n正在删除旧证书..."
        $store.Remove($existingCerts[0])
        $certInfo += "`n旧证书已删除`n"
    }
    
    # 8. 导入新证书
    $store.Add($cert)
    $store.Close()
    
    $certInfo += "正在导入新证书..."
    
    # 9. 成功弹窗
    $wshell.Popup("✅ 证书导入成功！`n`n$certInfo`n证书已添加到：本地计算机\受信任的根证书颁发机构", 0, "导入成功", 0 + 64)
    
} catch {
    # 10. 错误弹窗
    $errorMsg = "导入失败！`n`n错误信息：$_`n`n证书路径：$certFilePath"
    if ($certInfo) {
        $errorMsg = "导入失败！`n`n$certInfo`n错误信息：$_"
    }
    $wshell.Popup($errorMsg, 0, "导入错误", 0 + 16)
    exit 1
}
