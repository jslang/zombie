import generic

STYLE_MAP     ={  # Mapping for docx to intermediate styles
		'i'         : 'emphasis',
		'Emphasis'  : 'emphasis',
		'b'         : 'strong',
		'Strong'    : 'strong',
		'u'         : 'underline',
		'vanish'    : 'hidden',
		'webHidden' : 'hidden',
		'strike'    : 'strikethrough',
		'vertAlign' : 'superscript',
		}
PARSED_FILES = {} # Cache for parsed xml objects
LIST_STYLES  = {} # Cache for list styles
SYMBOL_MAP   = {  # Unicode (key) to Symbol font (value) mapping
	'0039': '39', '0038': '38', '0035': '35', '0034': '34', '0037': '37',
	'0036': '36', '0031': '31', '0030': '30', '0033': '33', '0032': '32',
	'007B': '7B', '21B5': 'BF', 'F6D9': 'D3', '002B': '2B', '2321': 'F5',
	'2320': 'F3', '002C': '2C', '2329': 'E1', '03D2': 'A1', '03D1': '4A',
	'2248': 'BB', '03D6': '76', '03D5': '6A', '002F': '2F', '00F7': 'B8',
	'2297': 'C4', '002E': '2E', '232A': 'F1', '2295': 'C5', '00AC': 'D8',
	'2260': 'B9', '007C': '7C', '221A': 'D6', 'F8E7': 'BE', '003E': '3E',
	'003D': '3D', '003F': '3F', '003A': '3A', 'F6DA': 'D2', '003C': '3C',
	'003B': '3B', '03A8': '59', '03A9': '57', '03BD': '6E', '03A4': '54',
	'03A5': '55', '03A6': '46', '03A7': '43', '03A0': '50', '03A1': '52',
	'03A3': '53', '0021': '21', '03BA': '6B', '2033': 'B2', 'F8FD': 'FD',
	'F8FE': 'FE', '0025': '25', 'F8FA': 'FA', 'F8FB': 'FB', 'F8FC': 'FC',
	'2666': 'A8', '2665': 'A9', '2663': 'A7', '2660': 'AA', '221D': 'B5',
	'2261': 'BA', '2264': 'A3', '2265': 'B3', '2282': 'CC', '2283': 'C9',
	'2284': 'CB', '2286': 'CD', '2287': 'CA', '00B0': 'B0', 'F8F9': 'F9',
	'F8F4': 'EF', 'F8F5': 'F4', 'F8F6': 'F6', 'F8F7': 'F7', 'F8F0': 'EB',
	'F8F1': 'EC', 'F8F2': 'ED', 'F8F3': 'EE', '2044': 'A4', 'F6DB': 'D4',
	'22A5': '5E', '0192': 'A6', '005B': '5B', '25CA': 'E0', '005F': '5F',
	'005D': '5D', 'F8E5': '60', '03C8': '79', '03C9': '77', '007D': '7D',
	'03C2': '56', '03C3': '73', '03C0': '70', '03C1': '72', '03C6': '66',
	'03C7': '63', '03C4': '74', '03C5': '75', '2190': 'AC', '2191': 'AD',
	'2192': 'AE', '2193': 'AF', '2194': 'AB', '221E': 'A5', '00A0': '20',
	'2200': '22', '2202': 'B6', '2203': '24', '2205': 'C6', '2206': '44',
	'2207': 'D1', '2208': 'CE', '2209': 'CF', '2118': 'C3', '2111': 'C1',
	'222A': 'C8', 'F8E6': 'BD', 'F8E9': 'E3', '211C': 'C2', 'F8EB': 'E6',
	'F8E8': 'E2', '220B': '27', '00B1': 'B1', '220F': 'D5', '00B5': '6D',
	'2212': '2D', '2211': 'E5', '2217': '2A', '2215': 'A4', '2026': 'BC',
	'F8ED': 'E8', '2022': 'B7', '22C5': 'D7', '0394': '44', '0395': '45',
	'0396': '5A', '0397': '48', '0391': '41', '0392': '42', '0393': '47',
	'0398': '51', '0399': '49', '2245': '40', '223C': '7E', '20AC': 'A0',
	'F8EE': 'E9', '2135': 'C0', '21D3': 'DF', '21D2': 'DE', '2228': 'DA',
	'2229': 'C7', '2227': 'D9', '2220': 'D0', '03B5': '65', '03B4': '64',
	'03B7': '68', '03B6': '7A', '03B1': '61', '2032': 'A2', '03B3': '67',
	'03B2': '62', '03B9': '69', '03B8': '71', '03BE': '78', '0023': '23',
	'0020': '20', '03BF': '6F', '0026': '26', 'F8EC': 'E7', '03BC': '6D',
	'03BB': '6C', '0028': '28', '0029': '29', '21D4': 'DB', 'F8F8': 'F8',
	'222B': 'F2', '00D7': 'B4', '21D1': 'DD', '2126': '57', '21D0': 'DC',
	'2234': '5C', '039D': '4E', '039E': '58', '039F': '4F', 'F8EF': 'EA',
	'F8EA': 'E4', '039A': '4B', '039B': '4C', '039C': '4D'}

def unicode_to_symbol(key):
	"""Returns the symbol code for unicode value key"""
	import docx
	map   = docx.SYMBOL_MAP
	try            : value = map[key]
	except KeyError: value = map.values()[-1]
	return int(value, 16)

def symbol_to_unicode(key):
	"""Returns the unicode for symbol value key"""
	import docx
	from time import time
	map   = dict([(v,k) for k,v in docx.SYMBOL_MAP.items()])
	try            : value = map[key]
	except KeyError: value = map.values()[-1]
	return int(value, 16)


class Docx(generic.Generic):
	def valid_in(self, input):
		""" Determines if the passed file is accepted by this module for input """
		#Guess by filename
		if not getattr(input, "read", False):
			valid_ext = ('docx',)
			ext       = input.split('.').pop().lower()
			return (ext in valid_ext)
		#Guess by file-like object
		else:
			head = input.read(4)
			input.seek(0)
			return head == "PK\x03\x04"
	
	def get_intermediate(self, input):
		""" Given an input, returns the intermediate representation of the file """
		from zipfile import ZipFile
		import docx
		import os
		
		if not getattr(input, "read", False): input = file(input)
		input = ZipFile(input, 'r')
		files = get_input_data(input)
		input.close()
		
		docx.PARSED_FILES = dict()
		document          = parse_file('word/document.xml', files).firstChild
		intermediate      = get_zombie(document, files)
		
		return intermediate


def parse_file(file, files, nocache=False):
	"""
	Parse file returns minidom object parsed from file in files.  Uses a 
	cache that can be disabled by passing True for argument nocache
	"""
	if file not in PARSED_FILES or nocache:
		from xml.dom import minidom
		PARSED_FILES[file] = minidom.parseString(files[file])
	return PARSED_FILES[file]

def get_input_data(input):
	""" 
	Read all archives in input into a dictionary with filenames as keys and
	return
	"""
	filenames = input.namelist()
	files     = [input.read(f) for f in filenames]
	files     = dict(zip(filenames, files))
	return files

def block_compact(element):
	import intermediate
	from modules import compact 
	
	test = lambda x,y: isinstance(x, intermediate.Text) and isinstance(y, intermediate.Text)
	element.children = compact(element.children, test)
	
	def test(x,y):
		if not isinstance(x, intermediate.Special): return False
		if not isinstance(y, intermediate.Special): return False
		if x.modifiers.symmetric_difference(y.modifiers): return False
		return True
	
	element.children = compact(element.children, test)
	return element

def get_zombietype(element):
	""" Returns the type (String) of element passed """
	testmap = {
		'Document'      : is_document,
		'Body'          : is_body,
		'Text'          : is_text,
		'Heading'       : is_heading,
		'Paragraph'     : is_paragraph,
		'List'          : is_list,
		'Table'         : is_table,
		'TableRow'      : is_tablerow,
		'TableHeader'   : is_tableheader,
		'TableCell'     : is_tablecell,
		'Image'         : is_image,
		'Link'          : is_link,
		'NewLine'       : is_newline,
		'SpecialRun'    : is_specialrun,
		'Run'           : is_run,
		'Insert'        : is_insert,
		}
	
	for type,func in testmap.items():
		if func(element): return type
	
	return 'Unknown'

def is_document(element):
	""" Test if element is document """
	if not element: return False
	return element.localName == 'document'

def is_body(element):
	""" Test if element is body """
	if not element: return False
	return element.localName == 'body'

def is_text(element):
	""" Test if element is text """
	if not element: return False
	return element.localName == 't' or element.localName == 'sym'

def is_heading(element):
	""" Test if element is heading """
	if not element: return False
	if element.localName == 'p':
		style = get_paragraph_style(element)
		style = style[:-1]
		return style == 'Heading'
	else:
		return False

def is_paragraph(element):
	""" Test if element is paragraph """
	if not element: return False
	return element.localName == 'p' and not is_heading(element)

def is_newline(element):
	""" Test for line break """
	if not element: return False
	return element.localName == 'br'

def is_specialrun(element):
	""" Test if this is a special run, one that is somehow different aside from
	data """
	if not element: return False
	if element.localName == 'r':
		try:
			#Check rPr children
			modifiers = element.getElementsByTagName('w:rPr').pop().childNodes
			modifiers = [mod for mod in modifiers if mod.localName in STYLE_MAP]
			if modifiers: return True
			
			#Check rStyle values
			modifiers = element.getElementsByTagName('w:rStyle')
			modifiers = [mod.getAttribute('w:val') for mod in modifiers]
			modifiers = [mod for mod in modifiers if mod in STYLE_MAP]
			if modifiers: return True
		except (AttributeError, IndexError): pass
	return False

def is_run(element):
	""" Test if element is a run """
	if not element: return False
	return element.localName == 'r' and not is_specialrun(element)

def is_insert(element):
	if not element: return False
	return element.localName == 'ins'

def is_list(element):
	""" Test for list """
	if not element: return False
	if element.localName == 'p':
		pPrs     = [child for child in element.childNodes if child.localName == 'pPr']
		return True in ['numPr' in [child.localName for child in pPr.childNodes] for pPr in pPrs]
	else: return False

def is_table(element):
	""" Test for table """
	if not element: return False
	return element.localName == 'tbl'

def is_tablerow(element):
	""" Test for tablerow """
	if not element: return False
	return element.localName == 'tr'

def is_tableheader(element):
	""" Test for tableheader """
	if not element: return False
	shd    = element.getElementsByTagName('w:shd')
	if not shd: return False
	
	ignore = set(('auto', 'FFFFFF', 'ffffff'))
	shaded = shd.pop().getAttribute('w:fill') not in ignore
	return is_tablecell(element) and shaded

def is_tablecell(element):
	""" Test for tablecell """
	if not element: return False
	return element.localName == 'tc'

def is_image(element):
	""" Test for image """
	if not element: return False
	elements = set(('pict', 'drawing',))
	return element.localName in elements

def is_link(element):
	""" Test for link """
	if not element: return False
	return element.localName == 'hyperlink'

def get_zombie(element, files):
	""" Get all elements conversion to intermediate"""
	import docx
	type = get_zombietype(element).lower()
	try                  : func = getattr(docx, 'get_zombie_' + type)
	except AttributeError: func = get_zombie_unknown
	return func(element, files)

def get_zombie_unknown(element, files):
	""" Handle unknown elements """
	import intermediate	
	ignore = set((
			'rPr', 'rStyle', 'pStyle', 'sectPr', 
			'pPr', 'tblPr', 'tcPr', 'proofErr',
			'trPr', 'smartTagPr', 'tblGrid',
			'commentRangeStart', 'commentRangeEnd',
			'commentReference', 'lastRenderedPageBreak',
			'del',
	))
	if element.localName in ignore:
		return None
	if element.localName == 'smartTag':
		return [get_zombie(child, files) for child in element.childNodes]
	else:
		return intermediate.Comment('Unknown element ' + element.localName)

def get_zombie_document(element, files):
	""" Return an intermediate version of the input document """
	import intermediate
	from datetime import datetime
	
	title    = get_title(files)
	author   = get_author(files)
	created  = datetime.now()
	
	document = intermediate.Document(title=title, author=author, created=created)
	
	for child in element.childNodes: 
		document.append(get_zombie(child, files))
	
	return document

def get_zombie_body(element, files):
	""" Get body elements """
	import intermediate
	body = intermediate.Body()
	
	for child in element.childNodes: 
		body.append(get_zombie(child, files))
	
	return body

def get_zombie_newline(element, files):
	import intermediate
	return intermediate.Text("\n")

def get_zombie_specialrun(element, files):
	""" Special runs are set apart in some way from	neighboring runs. """
	import intermediate
	from modules import compact
	
	rPr       = element.getElementsByTagName('w:rPr').pop()
	rStyles   = element.getElementsByTagName('w:rStyle')
	styles    = [style.getAttribute('w:val') for style in rStyles]
	styles    = set([STYLE_MAP[style] for style in styles if style in STYLE_MAP])
	modifiers = [child.localName for child in rPr.childNodes]
	modifiers = set([STYLE_MAP[modifier] for modifier in modifiers if modifier in STYLE_MAP])
	modifiers = modifiers.union(styles)
	
	special           = intermediate.Special()
	special.modifiers = modifiers
	
	for child in element.childNodes:
		child = get_zombie(child, files)
		special.append(child)
	
	special.children = compact(special.children, lambda x,y: isinstance(x, intermediate.Text) and isinstance(y, intermediate.Text))
	return special

def get_zombie_run(element, files):
	""" Get run text containers """
	import intermediate
	from modules import compact
	
	#A run is a list of text elements, treat as such
	run = list()
	for child in element.childNodes:
		run.append(get_zombie(child, files))
	
	run = compact(run, lambda x,y: isinstance(x, intermediate.Text) and isinstance(y, intermediate.Text))
	return run

def get_zombie_insert(element, files):
	"""
	Handle Word track change element, insert
	"""
	insert = list()
	for child in element.childNodes:
		insert.append(get_zombie(child, files))
	return insert

def get_zombie_link(element, files):
	""" Get zombie intermediate for link element """
	import intermediate
	
	try:
		rId  = element.getAttribute('r:id')
		rels = parse_file('word/_rels/document.xml.rels', files)
		rels = rels.getElementsByTagName('Relationship')
		rel  = [rel for rel in rels if rel.getAttribute('Id') == rId].pop()
		href = rel.getAttribute('Target')
	except IndexError:
		href = 'http://'
	
	link = intermediate.Link(href=href)
	
	for child in element.childNodes:
		child = get_zombie(child, files)
		link.append(child)	
	return link

def get_zombie_text(element, files):
	""" Get zombie intermediate for text element """
	import intermediate
	
	if element.localName   == 't'  :
		text = intermediate.Text()
		for child in element.childNodes: text = text + child.data
	elif element.localName == 'sym':
		symbol = element.getAttribute('w:char')
		symbol = hex(int(symbol, 16) & 0x0FFF)[2:]
		code   = symbol_to_unicode(symbol)
		text   = intermediate.Text(unichr(code))
	
	return text

def get_zombie_paragraph(element, files):
	""" Retrieve paragraph elements in intermediate form """
	import intermediate
	
	paragraph = intermediate.Paragraph()	
	for child in element.childNodes:
		paragraph.append(get_zombie(child, files))
		
	return block_compact(paragraph)

def get_zombie_heading(element, files):
	""" Retrieve heading elements in intermediate form """
	import intermediate
	
	style   = get_paragraph_style(element)
	level   = style[-1:].isdigit() and int(style[-1:]) or 0
	heading = intermediate.Heading(level)
	for child in element.childNodes:
		heading.append(get_zombie(child, files))
	
	return block_compact(heading)

def get_zombie_table(element, files):
	""" Retrieve table element in intermediate form """
	import intermediate
	table = intermediate.Table()
	for child in element.childNodes: table.append(get_zombie(child, files))
	return table

def get_zombie_tablerow(element, files):
	""" Retrieve tablerow element in intermediate form """
	import intermediate
	tr = intermediate.TableRow()
	for child in element.childNodes: tr.append(get_zombie(child, files))
	return tr

def get_zombie_tableheader(element, files):
	""" Retrieve tablecell element in intermediate form """
	import intermediate	
	
	colspan = get_colspan(element)
	rowspan = get_rowspan(element)	
	scope   = guess_scope(element)
	
	th = intermediate.TableHeader(colspan=colspan, rowspan=rowspan, scope=scope)
	for child in element.childNodes: th.append(get_zombie(child, files))
	return th

def get_zombie_tablecell(element, files):
	""" Retrieve tablecell element in intermediate form """
	import intermediate
	
	colspan = get_colspan(element)
	rowspan = get_rowspan(element)
	
	tc = intermediate.TableCell(colspan=colspan, rowspan=rowspan)
	for child in element.childNodes: tc.append(get_zombie(child, files))
	return tc

def get_zombie_image(element, files):
	""" Retrieve image object in intermediate form """
	if element.localName == 'drawing': image = get_drawing_image(element, files)
	if element.localName == 'pict'   : image = get_pict_image(element, files)
	
	#Add a prefix to prevent asset overwriting
	if image != None:
		from md5 import md5
		prefix = md5(image.data).hexdigest()[:5]
		name   = image.name.split('.')
		name.insert(-1, prefix)
		image.name = '.'.join(name)
	
	return image

def get_zombie_list(element, files):
	""" Retrieve a list element """
	import intermediate
	
	orderedtypes = ('decimal', 'lowerLetter', 'lowerRoman')
	import time
	type, lvl  = get_list_properties(element, files)
	if type in orderedtypes: list = intermediate.OrderedList()
	else                   : list = intermediate.UnorderedList()
	visited = []
	
	while element and is_list(element):
		ctype, clvl = get_list_properties(element, files)
		if clvl > lvl : #sublist
			list.children[-1].append(get_zombie_list(element, files))
		if clvl == lvl: #list item
			list.append(get_zombie_list_item(element, files))
		if clvl < lvl : #parent list
			break
		
		visited.append(element)
		element = element.nextSibling
	
	for item in visited[1:]: item.parentNode.removeChild(item)	
	return list

def get_zombie_list_item(element, files):
	""" Retrieve a list item """
	import intermediate
	
	list_item = intermediate.ListItem()	
	for child in element.childNodes:
		list_item.append(get_zombie(child, files))
		
	return block_compact(list_item)

def guess_scope(element):
	""" Given a table cell element, will attempt to guess its scope as if it were
	a table header """
	try:
		#Case 1: Cell heading row and not first column - col
		previous_is_th    = is_tableheader(element.previousSibling)
		parent_next_is_tr = is_tablerow(element.parentNode.nextSibling)	
		if	previous_is_th and parent_next_is_tr: return 'col'
		
		#Case 2: Cell is only one in row
		next_is_tc     = is_tablecell(element.nextSibling)
		previous_is_tc = is_tablecell(element.previousSibling)
		if not next_is_tc and not previous_is_tc: return 'col'
		
		#Case 3: Cell to the right of this cell is not a table header - row
		next_is_th = is_tableheader(element.nextSibling)
		if not next_is_th: return 'row'
		
		#Case 4: In first row, first col, and cell below is not table header - col
		in_first_col = not (is_tableheader(element.previousSibling) or is_tablecell(element.previousSibling))
		below_is_th  = is_tablerow(element.parentNode.nextSibling) and is_tableheader(element.parentNode.nextSibling.childNodes[0])
		if in_first_col and not below_is_th: return 'col'
	except IndexError: pass
	return None

def get_colspan(cell):
	""" Discovers the colspan of the table cell passed to it. Return None if N/A """
	gridSpan = cell.getElementsByTagName('w:gridSpan')
	if gridSpan: colspan = gridSpan.pop().getAttribute('w:val')
	else       : colspan = None
	return colspan

def get_rowspan(cell):
	""" Discovers the rowspan of the table cell passed to it. Return None if N/A """
	vMerge = cell.getElementsByTagName('w:vMerge')
	if not vMerge: return None
	
	parentrow = cell.parentNode
	hposition = parentrow.childNodes.index(cell)
	nextrow   = parentrow.nextSibling
	counter   = 1
	
	while nextrow and is_tablerow(nextrow):
		try              : cellbelow = nextrow.childNodes[hposition]
		except IndexError: break
		
		vMerge = cellbelow.getElementsByTagName('w:vMerge')
		if vMerge:
			if vMerge.pop().getAttribute('w:val') == 'restart': break
			counter += 1
			nextrow.removeChild(cellbelow)
		else: 
			break
		nextrow   = nextrow.nextSibling
	
	if counter > 1: return str(counter)
	else          : return None

def get_relationship(id, files):
	""" Given an id, will return the relationship defined in word's rel xml file
	or None if not found """
	try:
		rels = parse_file('word/_rels/document.xml.rels', files)
		rels = rels.getElementsByTagName('Relationship')
		rel  = [rel for rel in rels if rel.getAttribute('Id') == id].pop(0)
	except IndexError: 
		rel = None
	return rel

def get_drawing_image(element, files):
	""" Extracts image information from drawing elements in word XML. """
	import intermediate
	import os
	try:#to find image data and name
		blip = element.getElementsByTagName('a:blip').pop(0)
		rId  = blip.getAttribute('r:embed')
		rel  = get_relationship(rId, files)
		file = rel.getAttribute('Target')
		imgdata = files['word/' + file]
		imgname = os.path.basename(file)
	except IndexError: return None
	
	try:#to discover dimensions
		pixels_per_emu = 12700 #Constant for word's default metric (emu)
		 
		extent = element.getElementsByTagName('wp:extent').pop(0)
		x,y    = extent.getAttribute('cx'), extent.getAttribute('cy')
		x,y    = int(x)/pixels_per_emu, int(y)/pixels_per_emu
	except IndexError: x,y = None,None
	
	image = intermediate.Image(data=imgdata, name=imgname, height=y, width=x)
	return image

def get_pict_image(element, files):
	""" Extracts image information from pict elements in word XML. """
	import intermediate
	import os
	
	try:#to find image data and name
		imagedata = element.getElementsByTagName('v:imagedata').pop(0)
		rId       = imagedata.getAttribute('r:id')
		title     = imagedata.getAttribute('o:title')
		rel       = get_relationship(rId, files)
		file      = rel.getAttribute('Target')
		imgdata = files['word/' + file]
		imgname = os.path.basename(file)
	except IndexError: return None
	
	try:#to find image dimensions
		import re
		shape  = element.getElementsByTagName('v:shape').pop(0)
		style  = shape.getAttribute('style')
		alt    = shape.getAttribute('alt')
		result = re.search('width:([\d]+)pt;height:([\d]+)pt', style)
		if result: x,y = result.groups()
		else     : x,y = None,None
	except IndexError:
		x,y = None, None
		alt = None
		
	img = intermediate.Image(data=imgdata, name=imgname, height=y, width=x, descr=alt)
	return img

def get_title(files):
	""" Discover and return the title of the input document """
	if not files: return None
	
	core  = parse_file('docProps/core.xml', files)
	title = core.getElementsByTagName('dc:title')
	title = (title and title.pop().firstChild) or None
	if title: title = title.data.strip()
	else    : title = ''
	return title

def get_author(files):
	""" Discover and return the author of the input document """
	if not files: return None
	
	core   = parse_file('docProps/core.xml', files)
	author = core.getElementsByTagName('dc:author')
	return author and author.pop().firstChild.data or ''

def get_paragraph_style(element):
	""" Retrieve, if present, the paragraph style of an element.  Returns empty
	string if not present"""
	pStyle = element.getElementsByTagName('w:pStyle')
	if pStyle: 
		pStyle = pStyle.pop(0)
		pStyle = pStyle.getAttribute('w:val')
		return pStyle
	else: return str()

def get_list_properties(element, files):
	""" Return properties of a list, including the type and level """
	try:
		#Retrieve numbering id and list level
		numbering = parse_file('word/numbering.xml', files)
		numId     = element.getElementsByTagName('w:numId').pop(0).getAttribute('w:val')
		ilvl      = element.getElementsByTagName('w:ilvl').pop(0).getAttribute('w:val')
		
		#Check cache for id and lvl hash
		hash = '%s:%s' % (numId, ilvl)
		if hash in LIST_STYLES: return LIST_STYLES[hash]
		
		#Get abstract numbering id
		nums      = numbering.getElementsByTagName('w:num')
		num       = [num for num in nums if num.getAttribute('w:numId') == numId].pop(0)
		absNumId  = num.getElementsByTagName('w:abstractNumId').pop(0).getAttribute('w:val')
		
		#Get level element
		absNums   = numbering.getElementsByTagName('w:abstractNum')
		absNum    = [absNum for absNum in absNums if absNum.getAttribute('w:abstractNumId') == absNumId].pop(0)
		lvls      = absNum.getElementsByTagName('w:lvl')
		lvl       = [lvl for lvl in lvls if lvl.getAttribute('w:ilvl') == ilvl].pop(0)
		
		#Get number format value from numFmt element
		type = lvl.getElementsByTagName('w:numFmt').pop(0).getAttribute('w:val')
		
		#Store in cache
		LIST_STYLES[hash] = (type,ilvl)		
	except IndexError:
		type = 'bullet'
		ilvl = 0
	return type,ilvl
	