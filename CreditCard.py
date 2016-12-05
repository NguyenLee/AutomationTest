class CreditCard(object):
    def __init__(self, cardNumber, cvv, expiredMonth, expiredYear):
        self.cardNumber = cardNumber
        self.cvv = cvv
        self.expiredMonth = expiredMonth
        self.expiredYear = expiredYear