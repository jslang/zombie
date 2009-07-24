class Generic:
	""" Conversion classname should be the title cased version of its parent module """
	def valid_in(self, input):
		""" Determines if the passed filename is accepted by this module for input 
		Should return True/False"""
		return False
	
	def valid_out(self, output):
		""" Determines if the passed filename is accepted by this module for output 
		Should return True/False """
		return False
	
	def get_intermediate(self, input):
		""" Given an input, returns the intermediate representation of the file 
		Returns an intermediate form of the document, see the intermediate module"""
		return None
	
	def get_output(self, intermediate):
		""" Creates the Output object from the intermediate object
		Returns an output object, see the output module"""
		return None
