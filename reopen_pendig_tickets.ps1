[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

# reopen_pending_tickets
#
# Add note to ticket with a date in this format dd/MM/yyyy
# ex: To do on 23/05/2020
#
# Modify to match your environment
$APIKey = 'api_key'
$strDomain = 'domain.freshservice.com'
$intDaysSearch = -100
# Ticket status https://api.freshservice.com/v2/#update_ticket_priority
# Open	    2
# Pending	3
# Resolved	4
# Closed	5
$intGroupID = 3
$strMail = 'test@example.com'
$SmtpServer = 'mail.example.com'
#######################################

$EncodedCredentials = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $APIKey,$null)))
$HTTPHeaders = @{}
$HTTPHeaders.Add('Authorization', ("Basic {0}" -f $EncodedCredentials))
$HTTPHeaders.Add('Content-Type', 'application/json')

# Recupero da quanti giorni fare la ricerca dei ticket
$objDataInizioRicerca = (Get-Date).AddDays($intDaysSearch)
$strDataInizioRicerca = $objDataInizioRicerca.ToString("yyyy-MM-dd")

### Statistiche
$intNumeroTotaleRichiesteAPI = 0
$intNumeroTotaleTicketAnalizzati = 0
$intNumeroTicketGruppo7 = 0
$intNumeroTicketCambiati = 0
$intNumeroTicketConData = 0
$log = "`n`nData " + (Get-Date).ToString() + "`n"
###

For ($i=1; $i -le 10000; $i++) {    
    
    $URL = "https://" + $strDomain +  "/api/v2/tickets?updated_since=" + $strDataInizioRicerca + "T02:00:00Z&page=$i"
    $result = Invoke-RestMethod -Method Get -Uri $URL -Headers $HTTPHeaders
    $intNumeroTotaleRichiesteAPI = $intNumeroTotaleRichiesteAPI + 1

    # Write-Host -ForegroundColor DarkGreen -BackgroundColor Yellow $result.tickets.Count

    # Conto il numero di ticket per statistica
    $intNumeroTotaleTicketAnalizzati = $intNumeroTotaleTicketAnalizzati + $result.tickets.Count
    # Esco dal ciclo se il numero di ticket recuperati è pari a 0
    if ($result.tickets.Count -eq 0) { break }
    
    foreach ($ticket in $result.tickets) {
        if ($ticket.status -eq $intGroupID) {
            Write-Host $ticket.id
            $log = $log + "---`nTicketID: " + $ticket.id + "`n"
            Write-Host $ticket.subject
            $log = $log + "Subject: " + $ticket.subject + "`n"

            # Statistiche
            $intNumeroTicketGruppo7 = $intNumeroTicketGruppo7 + 1

            $intTicketID = $ticket.id

            $URL2 = "https://" + $strDomain +  "/api/v2/tickets/" + $intTicketID + "?include=conversations"
            $result2 = Invoke-RestMethod -Method Get -Uri $URL2 -Headers $HTTPHeaders
            $intNumeroTotaleRichiesteAPI = $intNumeroTotaleRichiesteAPI + 3
                
            foreach ($t in $result2) {

                foreach ($conversation in $t.ticket.conversations) {
                    if ($conversation.body_text.Contains("/")) {
                        
                        $strForseData = $conversation.body_text
                        $array_strForseData = $strForseData.Split("/")
                        $length = $array_strForseData[0].length
                        $giorno = $array_strForseData[0].substring($length -2, 2)
                        $mese = $array_strForseData[1]
                        $anno = $array_strForseData[2].substring(0, 4)
        
                        $data = $giorno + "-" + $mese + "-" + $anno
        
                        try { $dateDataRiaperturaTicket = [datetime]::parseexact($data, 'dd-MM-yyyy', $null) }
                        catch { continue }
                        
                        $intNumeroTicketConData = $intNumeroTicketConData + 1

                        Write-Host -BackgroundColor Red "Note contains a date:" $dateDataRiaperturaTicket.ToString("dd-MM-yyyy")
                        $log = $log + "Note contains a date: " + $dateDataRiaperturaTicket.ToString("dd-MM-yyyy") + "`n"

                        if ((Get-Date) -ge $dateDataRiaperturaTicket) {
                            Write-Host -BackgroundColor Red "il ticket è da riaprire"

                            $UserAttributes = @{}
                            $UserAttributes.Add('status' , 2)
                            $UserAttributes = @{'helpdesk_ticket' = $UserAttributes}
                            $JSON = $UserAttributes | ConvertTo-Json 

                            Write-Host $JSON

                            # Uso le API v1 perchè con le v2 non riesco ad aggiornare il ticket
                            $URL3 = "https://" + $strDomain +  "/helpdesk/tickets/$intTicketID.json"
                            Write-Host $URL3
                            Invoke-RestMethod -Method Put -Uri $URL3 -Headers $HTTPHeaders -Body $JSON
                            $intNumeroTotaleRichiesteAPI = $intNumeroTotaleRichiesteAPI + 1
                            $intNumeroTicketCambiati = $intNumeroTicketCambiati + 1
                        }
                    }
                }
            }
        }
    }
}

$log = $log + "`nAPI requests:                   " + $intNumeroTotaleRichiesteAPI
$log = $log + "`nTickets analyzed:               " + $intNumeroTotaleTicketAnalizzati
$log = $log + "`nTickets in pending state:       " + $intNumeroTicketGruppo7
$log = $log + "`nTickets updated:                " + $intNumeroTicketCambiati
$log = $log + "`nTickets with date:              " + $intNumeroTicketConData
$log = $log + "`nTickets without date:           " + ($intNumeroTicketGruppo7 - $intNumeroTicketConData)

Write-Host $log

Send-MailMessage -From 'reopen_pending_tickets@no-reply.com' -To $strMail -Subject 'reopen_pending_tickets' -Body $log -SmtpServer $SmtpServer