# Create AD user with GUI
Deze powershell script met een GUI, kun je 1 of met CSV bulk import functie active directory users aanmaken waarbij een automatisch gegenereerd password per mail wordt verstuurd.
Maakt automatisch username aan op basis van Voor en Achternaam:
Jan Pet:
username = j.pet
mail =j.pet@mijndomain.nl

Jan met de Pet:
username = j.metdepet
mail = j.metdepet@mijndomain.nl

Features: 
- Password Policy
- Notify Email (alleen single user)
- UPN domain selectie
- OU Path selectie
- Account Preview (alleen single user)

Email settings kunnen ook in de GUI ingesteld worden:
- SMTP Server 
- SSL supported
- Email settings worden opgeslagen in een xml file, wachtwoord wordt encrypted opgeslagen

![image](https://github.com/user-attachments/assets/e6259e14-379f-4cd7-a1d7-a1477c246855)
