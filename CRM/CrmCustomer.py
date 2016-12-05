class CrmCustomer:
    def __init__(self, firstname, lastname, addressObj, addressType,
                        email = None, emailType = 0,
                        phone = None, phoneType = 0):
        self.firstname = firstname
        self.lastname = lastname
        self.addressObj = addressObj
        self.addressType = addressType
        self.email = email
        self.emailType = emailType
        self.phone = phone
        self.phoneType = phoneType