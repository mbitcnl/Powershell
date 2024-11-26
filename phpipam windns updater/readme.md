# phpIPAM Windows DNS Updater
Deze powershell script is gemaakt voor het bijwerken van DNS records in Windows Server op basis van de gegevens uit PHPIPAM, middels de phpIPAM API haalt het script IP-adressen en hostnames op uit PHPIPAM, voegt of werkt de DNS-records bij en controleert of de PTR-records correct zijn en gaat deze toevoegen indien deze er niet zijn.
Het script logt alle acties en fouten in een logbestand. 
