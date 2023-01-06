function getusers {
    param (
        $CN
    )
    $users = Get-Aduser -properties Name, PasswordNeverExpires, PasswordExpired, PasswordLastSet, EmailAddress -filter * | Where-Object { ($_.DistinguishedName -match $CN) -and ($_.EmailAddress -ne $null) -and ($_.PassWordLastSet -ne $null)}
    return $users
}

function treatment {
    param (
            $data
    )
    $today = (get-date)
    $maxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge

    foreach ($user in $data) {
        $Name = (Get-ADUser $user | ForEach-Object { $_.Name }) -split (" ")
        $passwordsetdate = $user.passwordLastSet
        $expiresOn = $passwordsetdate + $maxPasswordAge
        $daystoexpire = -((New-TimeSpan -Start $expiresOn -End $today).Days)
        $dayson = $false

        Switch ($daystoexpire) {
        "1" {
                $dayson = $true
            }
        "2" {
                $dayson = $true
            }
        "3" {
                $dayson = $true
            }
        "4" {
                $dayson = $true
            }
        "5" {
                $dayson = $true
            }
        "15" {
                $dayson = $true
             }
        "30" {
                $dayson = $true
             }
        default {}
         }
        if($dayson -or ($daystoexpire -lt 0)) {
            $prenom = $Name[1]
            $emailaddress = $user.emailaddress

            $MailFrom = "XXXXXXXXXX"
            $MailTo  = $emailaddress

            # Server Info
            $SmtpServer = "mail.humans.local"
            $SmtpPort = "25"

            # Message stuff
            $subjectexpire = "expire dans $daystoexpire"
            $messexpire = "expire dans <b>$daystoexpire</b>"
            $abs = $day
            if($daystoexpire -lt 0) {
                $abs = $daystoexpire * -1
                $subjectexpire = "a déjà expiré depuis $abs"
                $messexpire = "a déjà expiré depuis <b>$abs</b>"
            }   

            $subject = "Votre mot de passe Windows $subjectexpire jours" 
            $Message = New-Object System.Net.Mail.MailMessage $MailFrom,$MailTo
            $Message.IsBodyHTML = $true
            $Message.Subject = $subject
               
            $Message.Body = @"

            <p>Bonjour <b>$prenom</b>,</p>

            <p>Votre mot de passe de session Windows $messexpire jours.</p>

            <p>Voici la procédure pour le changer :</p>

            <p>
            - Si vous êtes en télétravail, veuillez lancer le VPN. Sinon veuillez vous rendre au 
            siège de votre entreprise.<br>
            - Appuyez sur Ctrl-Alt-Suppr de votre clavier et sélectionnez 'Changer le mot de passe'.
            </p>

            <p>
            Si vous rencontrez un problème, n'hésitez pas à envoyer un mail à <b>support@data-expertise.com</b> ou
            à nous appeler directement au <b>09 78 23 20 29</b> et appuyez sur le '2' pour rentrer en contact avec
            l'un de nos techniciens.
            </p>
                     
            <p>
            Bien cordialement,<br>
            Le support de Data-expertise
            </p
"@
            # Construct the SMTP client object, credentials, and send
            $Smtp = New-Object Net.Mail.SmtpClient($SmtpServer,$SmtpPort)
            $Smtp.Send($Message)
        }
    }
}

function main {
    param (
        $expiracydays,
        $CN
    )

    $data = getusers($CN)
    
    treatment($data)
}

#VARS
$CN = "OU=Admins,OU=Utilisateurs,DC=XXX,DC=XXX"
main($expiracydays, $CN)
