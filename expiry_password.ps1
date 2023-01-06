function getusers {
    param (
        $CN
    )
    #Recupere les utilisateurs actifs ayant le champ adresse-email remplie
    $users = Get-Aduser -properties Name, PasswordNeverExpires, PasswordExpired, PasswordLastSet, EmailAddress -filter * | Where-Object { ($_.DistinguishedName -match $CN) -and ($_.EmailAddress -ne $null) -and ($_.PassWordLastSet -ne $null)}
    #Retourne un tableau contenant les utilisateurs
    return $users
}

function treatment {
    param (
            $data
    )
    #Je stocke la date actuel
    $today = (get-date)
    #Je stocke l'âge maximal du mot de passe définit dans la GPO 'Default Domain Policy'
    $maxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge

    #Je parcours mon tableau d'utilisateur
    foreach ($user in $data) {
        #Je stocke dans un tableau le nom et le prénom de l'utilisateur
        $Name = (Get-ADUser $user | ForEach-Object { $_.Name }) -split (" ")
        #Je recupere la date de la dernière modification du mot de passe de l'utilisateur
        $passwordsetdate = $user.passwordLastSet
        #Je calcule la date d'expiration du mot de passe
        $expiresOn = $passwordsetdate + $maxPasswordAge
        #Je convertie la date calculé en jours.
        $daystoexpire = -((New-TimeSpan -Start $expiresOn -End $today).Days)
        
        #J'initialise un boolean qui me servira à savoir si j'envoie ou non le mail
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

         # si la date d'expiration est égalle à 1,2,3,4,5,15,30 ou inférieur à 0, j'envoie un mail.
        if($dayson -or ($daystoexpire -lt 0)) {
            #Je stocke le prenom de l'utilisateur
            $prenom = $Name[1]
            #Je stocke l'adresse email de l'utilisateur
            $emailaddress = $user.emailaddress

            #Renseignez ici l'emetteur du mail
            $MailFrom = "XXXXXXXXX"
            #Renseignez ici le destinateur du mail ici l'utilisateur
            $MailTo  = $emailaddress

            # Serveur Info
            $SmtpServer = "XXXXXXX"
            $SmtpPort = "XX"

            #Je renseigne ici une partie du sujet du mail
            $subjectexpire = "expire dans $daystoexpire"
            #message utilisé dans le corps du mail
            $messexpire = "expire dans <b>$daystoexpire</b>"

            #si la date d'expiration est inferieur à 0, je change le sujet et un bout de message utilisé 
            if($daystoexpire -lt 0) {
                #vu que la date d'expiration est un chiffre/nombre negatif, je le convertie en nombre positif
                $absday = $daystoexpire * -1
                #je modifie la partie du sujet du mail
                $subjectexpire = "a déjà expiré depuis $absday"
                #je modifie un bout de message utilise dans le corps
                $messexpire = "a déjà expiré depuis <b>$absday</b>"
            }   

            # je créé le sujet du mail
            $subject = "Votre mot de passe Windows $subjectexpire jours" 
            #je créé le corps du mail
            $Message = New-Object System.Net.Mail.MailMessage $MailFrom,$MailTo
            #j'active l'HTML dans le corps du mail
            $Message.IsBodyHTML = $true
            #Je renseigne le sujet du mail
            $Message.Subject = $subject
            
            #Je créé le corps du mail   
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
            #Je créé l'objet me permettant d'envoyer le mail
            $Smtp = New-Object Net.Mail.SmtpClient($SmtpServer,$SmtpPort)
            #J'envoie le mail à l'utilisateur
            $Smtp.Send($Message)
        }
    }
}

function main {
    param (
        $CN
    )
    #Recupere les utilisateurs actifs ayant le champ adresse-email remplie
    $data = getusers($CN)
    
    treatment($data)
}

#VAR
#Insérez le chemin où se trouve les utilisateurs
$CN = "OU=Admins,OU=Utilisateurs,DC=XXXX,DC=XXXXX"
main($CN) treatment($data)
