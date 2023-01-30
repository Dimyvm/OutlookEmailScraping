import win32com.client


def SendErrorMail(title, error):
    outlook = win32com.client.dynamic.Dispatch("Outlook.Application")
    olMailItem = 0x0
    mail = outlook.CreateItem(olMailItem)
    emailAdress = mail.Session.CurrentUser.Address  # get mailadress current user
    # mail.To = 'dvanmulders@trevi-env.com'
    mail.To = emailAdress
    mail.Subject = 'ERROR - Siemens email scraping'
    mail.BodyFormat = 2  # olFormatHTML

    mail.HTMLBody = f'''<h2>Bij het uitvoeren van het script op is een error vastgesteld.</h2>
    <h3>Hieronder vindt u een overzicht van de error</h3>
    <h4>{title}</h4>
    <p>{error}</p>'''  # this field is optional
    mail.display()
    mail.Send()
