class HotelAddress(object):
	def __init__(self, name, city, isOutsideUS = False, roomNumber = 0, stateCode = None, address1 = None, address2 = None, zip = None):
		self.name = name
		self.city = city
		self.isOutsideUS = isOutsideUS
		self.roomNumber = roomNumber
		self.stateCode = stateCode
		self.address1 = address1
		self.address2 = address2
		self.zip = zip
		if not (self.isOutsideUS) :
			if (self.stateCode is None or self.zip is None) :
				raise Exception('This hotel in US, so, it must have state & zip')

