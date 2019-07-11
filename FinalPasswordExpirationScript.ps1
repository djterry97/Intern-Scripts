$AD = Get-ADUser -filter { Enabled -eq $True -and PasswordNeverExpires -eq $False } -Properties 'Name', 'msDS-UserPasswordExpiryTimeComputed', 'EmailAddress'
$Data = $AD | Select-Object 'Name', 'msDS-UserPasswordExpiryTimeComputed', 'EmailAddress'
$CurDate = Get-Date
$Data | ForEach-Object {
  $daysToExpiration = ( [datetime]::FromFileTime( $_.'msDS-UserPasswordExpiryTimeComputed' ) - ( $CurDate ) ).days  
  $Days = switch ( $daysToExpiration )
  {
    3 { 3 } 7 { 7 } 14 { 14 }    
    default { $null }
  }
  if ( $Days -ne $null ) { 
    Write-Host "$( $_.Name ) Expiring in $Days days"    
    $Email = @{
      To = $_.EmailAddress
      From = 
      Subject = 'Account Password Expiring Soon'
      Body = $_.Name + ",`n" + "Your password will expire in " + $Days + " days.`n" + $_.EmailAddress
      SMTPServer = ""
    }
    Send-MailMessage @Email    
    $Days = $null
  }
}
