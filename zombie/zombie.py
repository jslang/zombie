class Zombie:
	""" The Zombie class is responsible for finding modules for i/o conversion
	and storing and providing the data that is returned by the input and output
	modules."""
	def __init__(self, input, output, imod=None, omod=None, title=None):
		"""
		When passing input and output file as strings, input and output are the
		only required fields.  However, input may also be passed as a file-like
		object as well, and in this case the imod argument _must_ be specified
		or a TypeError exception will be thrown..
		"""
		from os import path
		
		self.title = title
		
		if not getattr(input, "read", False):
			self.input_file = path.realpath(path.expanduser(input))
		else:
			self.input_file = input
		self.output_file = path.realpath(path.expanduser(output))
		
		self.imod = self.__get_imod(imod)
		self.omod = self.__get_omod(omod)
		
		self.intermediate = None
		self.output       = None
	
	def convert(self):
		""" Walks through the stages of conversion, input to intermediate, 
		intermediate to output.  Returns True/False to indicate failure."""
		#Don't even try if we have no modules to use
		if not self.imod or not self.omod: return False
		
		self.intermediate = self.imod.get_intermediate(self.input_file)
		
		#Override document title if passed as argument to init
		if self.intermediate and self.title : self.intermediate.title = self.title
		
		self.output = self.omod.get_output(self.intermediate)
		
		#Check for failure
		if self.intermediate and self.output: return True
		else                                : return False
		
	def finalize(self):
		"""
		Will call the current job output object's method final in a seperate
		thread and return that thread.
		"""
		from threading import Thread
		output_thread = Thread(target=self.output.final, args=[self.output_file,])
		output_thread.start()
		return output_thread
	
	def __get_imod(self, imod=None):
		""" Get the specified input module, or guess if not given """
		if imod: mod = self.__get_mod(imod, input=True) 
		else   : mod = self.__get_mod(input=True)
		return mod
	
	def __get_omod(self, omod=None):
		""" Get the specified output module, or guess if not given """
		if omod: mod = self.__get_mod(omod, output=True) 
		else   : mod = self.__get_mod(output=True)
		return mod
	
	def __get_mod(self, module=None, input=False, output=False):
		""" Get the module specified, or guess if not given """
		import modules
		if module: return import_module(module)
		else     :
			modules = [import_module(module) for module in modules.__all__]
			modules = [module for module in modules if self.__test_module(module, input, output)]
			#Return the first valid module
			if modules: return modules.pop(0)
		
		return False
	
	def __test_module(self, module, input, output):
		""" Determine if the module given is valid for either input and/or output """
		
		valid_in, valid_out = False, False
		if input  and module.valid_in(self.input_file)  : valid_in = True
		if output and module.valid_out(self.output_file): valid_out = True
		
		if input and output: return valid_in and valid_out
		if input           : return valid_in
		if output          : return valid_out
		return False


def import_module(modulename):
	""" Import a zombie module given by it's name """
	import modules
	import sys
	name   = "%s.%s" % (modules.__name__, modulename)
	__import__(name)
	module = sys.modules[name]
	return getattr(module, to_classname(module.__name__))()

def to_classname(modulename):
	""" Convert a module's name to what it's contained conversion class should be named"""
	modulename = modulename.split('.')
	return modulename.pop().title()