import threading

class OutputThread (threading.Thread):
	""" Class to handle threading for output writing """
	def __init__(self, procedure, *args, **kwargs):
		self.procedure = procedure
		self.args      = args
		self.kwargs    = kwargs
		super(OutputThread, self).__init__()
	
	def run(self):
		self.procedure(*self.args, **self.kwargs)

class Output:
	""" The Output class is responsible for describing the converted document
	and providing any interfaces necessary for the output of that document"""
	def __init__(self, data, assets=None, meta={}):
		self.data   = data
		self.assets = assets
		self.meta   = meta
		
	def __repr__(self):
		return self.data
	
	def final(self, out):
		"""The final method is responsible for providing the final document
		ie. Writing to disk, archiving, etc"""
		from os import path
		outdir  = path.dirname(out)
		outname = path.basename(out)
		
		filename = path.join(outdir, outname)
		f        = file(filename, 'w')
		f.write(self.data)
		f.close()
		
		if not self.assets: return
		for name,data in self.assets.items():
			filename = path.join(outdir, name)
			f        = file(filename, 'w')
			f.write(data)
			f.close()


class Asset:
	"""The Asset class provides a datatype for objects that are required by the
	document, but are not part of the document.  Essentially binary data."""
	def __init__(self, data, name):
		self.data = data
		self.name = name