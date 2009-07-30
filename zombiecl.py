#!/usr/bin/python
import sys

#Run as command line app
def main():
	from zombie import Zombie
	from zombie import modules
	from optparse import OptionParser
	
	#Check python version
	required = (2,4)
	version  = sys.version_info[:2]
	if version < required:
		required = '.'.join(map(lambda x: str(x), required))
		print 'Zombie requires at least Python ', required
		return 0
	
	installed = modules.__all__
	usage = """usage: %prog INPUT OUTPUT [options]"""
	descr = """Zombie converts documents from one arbitrary format to another."""
	
	parser = OptionParser(usage=usage, description=descr)
	parser.add_option('-i', dest='imodule', help='Override auto detection for input module', type='string')
	parser.add_option('-o', dest='omodule', help='Override auto detection for output module', type='string')
	parser.add_option('-e', dest='encoding', help='Specify encoding for output module, not all modules use this.', type='string')
	parser.add_option('-v', dest='verbose', help='Enable verbose output', action='store_true')
	parser.add_option('-t', dest='title', help='Explicitly set title for conversion', type='string')
	parser.add_option('-l', dest='list', help='List available modules', action="store_true")
	options, args = parser.parse_args()
	
	if len(args) < 2 or options.list:
		if options.list:
			print 'Modules available to Zombie:'
			for module in modules.__all__: print "\t* " + module
		else: parser.print_help()
		return 0
	
	args = {
		'input'   : args[0],
		'output'  : args[1],
		'omod'    : options.omodule,
		'imod'    : options.imodule,
		'title'   : options.title,
		'encoding': options.encoding,
		'verbose' : options.verbose,
	}
	
	job = Zombie(
		input  = args['input'], 
		output = args['output'], 
		imod   = args['imod'],	
		omod   = args['omod'], 
		title  = args['title']
	)
	if job.omod:
		job.omod.VERBOSE = args['verbose'] or False
		job.omod.CHARSET = args['encoding'] or 'utf-8'
	
	if job.convert():
		output_thread = job.finalize()
		output_thread.join() #Wait for output thread to rejoin
	else:
		errors = ['Conversion failed for the following reasons:',]
		if not job.imod:
			errors.append('No suitable input module found')
		if not job.omod: 
			errors.append('No suitable output module found')
		if not job.intermediate: 
			errors.append('Input module did not produce a usable intermediate object')
		if not job.output: 
			errors.append('Output module did not produce a usable output object')
		print '\n\t'.join(errors)
	return 0
	
sys.exit(main())