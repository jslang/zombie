## Description ##

Zombie is an application that is meant to provide an easily extensible tool to 
convert a document from one format to another. Actual conversion takes place in
modules that can be written to interface with the internal representation of 
document data. These modules are responsible for conversion to and from this 
internal representation.

Zombie itself is responsible for the handling of input and output, as well as
loading the appropriate modules for conversion. Additionally, it is to monitor 
for any errors that may occur and inform the user appropriately.

## Usage ##

Zombie can be used via the command-line as a standalone program or as part of a
python program as a python module.

### Command Line Usage ###

General usage of zombie with an INPUT file and OUTPUT file would be:

`python zombiecl.py INPUT OUTPUT`

Zombie will attempt to autodetect the appropriate modules to use for input and 
output.  It will use the first module it finds, which may not be the module you
want.  In this case, you can also specify the module names:

`python zombiecl.py INPUT OUTPUT [-i input_module] [-o output_module]`

You can specify the input module, output module, or both.  Refer to individual 
module documentation for what names to use for these arguments.

A more detailed help description can be found by running:

`python zombiecl.py --help`

### Module Usage ###

Importing the zombie module should be sufficient to use it in another python 
application.  This can be achieved by doing the following:

	from zombie import Zombie
	job = Zombie(input=INPUT, output=OUTPUT)
	if job.convert(): job.finalize()

The converted document can be accessed without finalizing it.  The Output object
defined by Zombie.output contains all the data from the conversion.

	from zombie import Zombie
	job = Zombie(input=INPUT, output=OUTPUT)
	if job.convert():
	    document = job.output.data     #The actual data for the converted document.
	    assets   = job.output.assets   #Any ancillary assets needed by the document.

In this way you can do whatever you wish with the outputted document.

### Miscellaneous ###

Some output modules support the optional encoding argument.  This can specify 
which encoding to use for output, such as utf-8, ascii, etc.  As a module this 
is achieved by the following:

	job = Zombie(input=INPUT, output=OUTPUT)
	job.omod.CHARSET = 'utf-8'
	if job.convert(): job.finalize()

For command line usage, refer to the command line help for the appropriate 
option to pass.
