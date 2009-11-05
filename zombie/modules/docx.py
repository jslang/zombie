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
		'vertAlign' : {
			'superscript' : 'superscript',
			'subscript'   : 'subscript',
			},
		}
PARSED_FILES = {} # Cache for parsed xml objects
LIST_STYLES  = {} # Cache for list styles
SYMBOL_MAP   = {  # Mapping for symbol to Unicode
	0xBD:0x23D0, 0xBE:0x23AF, 0xBF:0x21B5, 0xBA:0x2261, 0xBB:0x2248,
	0xBC:0x2026, 0xFB:0x23A6, 0x5E:0x22A5, 0x5D:0x005D, 0x5F:0x005F,
	0x5A:0x0396, 0x5C:0x2234, 0x5B:0x005B, 0x24:0x2203, 0x25:0x0025,
	0x26:0x0026, 0x27:0x220D, 0x20:0x0020, 0x21:0x0021, 0x22:0x2200,
	0x23:0x0023, 0x28:0x0028, 0x29:0x0029, 0xF1:0x3009, 0xF2:0x222B,
	0x2D:0x2212, 0x2E:0x002E, 0x2F:0x002F, 0xF3:0x2320, 0x2A:0x2217,
	0x2B:0x002B, 0x2C:0x002C, 0xF4:0x23AE, 0xF5:0x2321, 0x59:0x03A8,
	0x58:0x039E, 0xFA:0x23A5, 0x55:0x03A5, 0x54:0x03A4, 0x57:0x03A9,
	0x56:0x03C2, 0x51:0x0398, 0x50:0x03A0, 0x53:0x03A3, 0x52:0x03A1,
	0xB4:0x00D7, 0xB5:0x221D, 0xB6:0x2202, 0xB7:0x2022, 0xB0:0x00B0,
	0xB1:0x00B1, 0xB2:0x2033, 0xB3:0x2265, 0xC0:0x2135, 0xB8:0x00F7,
	0xB9:0x2260, 0xFE:0x23AD, 0x3C:0x003C, 0x3B:0x003B, 0x3A:0x003A,
	0x3F:0x003F, 0x3E:0x003E, 0x3D:0x003D, 0xE9:0x23A1, 0xE8:0x239D,
	0xE5:0x2211, 0xE4:0x0021, 0xE7:0x239C, 0xE6:0x239B, 0xE1:0x3008,
	0xE0:0x22C4, 0xE3:0x0000, 0xE2:0x0000, 0xF6:0x239E, 0xEE:0x23A9,
	0xED:0x23A8, 0xEF:0x23AA, 0xEA:0x23A2, 0xEC:0x23A7, 0xEB:0x23A3,
	0x39:0x0039, 0x38:0x0038, 0x33:0x0033, 0x32:0x0032, 0x31:0x0031,
	0x30:0x0030, 0x37:0x0037, 0x36:0x0036, 0x35:0x0035, 0x34:0x0034,
	0xFC:0x23AB, 0x60:0xF8E5, 0x61:0x03B1, 0x62:0x03B2, 0x63:0x03C7,
	0x64:0x03B4, 0x65:0x03B5, 0x66:0x03C6, 0x67:0x03B3, 0x68:0x03B7,
	0x69:0x03B9, 0xFD:0x23AC, 0xC9:0x2283, 0xC8:0x222A, 0xC3:0x2118,
	0xC2:0x211C, 0xC1:0x2111, 0x78:0x03BE, 0xC7:0x2229, 0xC6:0x2205,
	0xC5:0x2295, 0xC4:0x2297, 0xCC:0x2282, 0xCB:0x2284, 0xCA:0x2287,
	0xCF:0x2209, 0xCE:0x2208, 0xCD:0x2286, 0xF0:0xF8FF, 0x6A:0x03D5,
	0x6B:0x03BA, 0x6C:0x03BB, 0x6D:0x03BC, 0x6E:0x03BD, 0x6F:0x03BF,
	0xF7:0x239F, 0xF8:0x23A0, 0xF9:0x23A4, 0xDF:0x21D3, 0xDD:0x21D1,
	0xDE:0x21D2, 0xDB:0x21D4, 0xDC:0x21D0, 0xDA:0x2228, 0x7E:0x223C,
	0x7D:0x007D, 0x7C:0x007C, 0x7B:0x007B, 0x48:0x0397, 0x49:0x0399,
	0x46:0x03A6, 0x47:0x0393, 0x44:0x0394, 0x45:0x0395, 0x42:0x0392,
	0x43:0x03A7, 0x40:0x2245, 0x41:0x0391, 0xA1:0x03D2, 0xA0:0x20AC,
	0xA3:0x2264, 0xA2:0x2032, 0xA5:0x221E, 0xA4:0x2044, 0xA7:0x2663,
	0xA6:0x0192, 0xA9:0x2665, 0xA8:0x2666, 0xAA:0x2660, 0xAC:0x2190,
	0xAB:0x2194, 0xAE:0x2192, 0xAD:0x2191, 0xAF:0x2193, 0x77:0x03C9,
	0x76:0x03D6, 0x75:0x03C5, 0x74:0x03C4, 0x73:0x03C3, 0x72:0x03C1,
	0x71:0x03B8, 0x70:0x03C0, 0x4F:0x039F, 0x4D:0x039C, 0x4E:0x039D,
	0x4B:0x039A, 0x4C:0x039B, 0x79:0x03C8, 0x4A:0x03D1, 0x7A:0x03B6,
	0xD8:0x00AC, 0xD9:0x2227, 0xD6:0x221A, 0xD7:0x22C5, 0xD4:0x2122,
	0xD5:0x220F, 0xD2:0x00AE, 0xD3:0x00A9, 0xD0:0x2220, 0xD1:0x2207,
}
WINGDINGS_MAP = { # Mapping for Wingding to Unicode
	0xBD:0x0000, 0xBE:0x0000, 0xBF:0x0000, 0xBA:0x0000, 0xBB:0x0000,
	0xBC:0x0000, 0xFB:0x2717, 0x5E:0x2648, 0x5D:0x2638, 0x5F:0x2649,
	0x5A:0x262A, 0x5C:0x0950, 0x5B:0x262F, 0x24:0x0000, 0x25:0x0000,
	0x26:0x0000, 0x27:0x0000, 0x20:0x0020, 0x21:0x270E, 0x22:0x2702,
	0x23:0x2701, 0x28:0x260E, 0x29:0x2706, 0xA4:0x2609, 0x87:0x2466,
	0xF1:0x21E7, 0xF2:0x21E9, 0x2D:0x0000, 0x2E:0x0000, 0x2F:0x0000,
	0xF3:0x21D4, 0x2A:0x2709, 0x2B:0x2709, 0x2C:0x0000, 0xF4:0x21D5,
	0xF5:0x21D6, 0x59:0x2721, 0x58:0x2720, 0xFA:0x0000, 0x55:0x271E,
	0x54:0x2744, 0x57:0x271D, 0x56:0x271E, 0x51:0x2708, 0x50:0x0000,
	0x53:0x0000, 0x52:0x263C, 0xB4:0xFFFD, 0xB5:0x272A, 0xB6:0x2730,
	0xB7:0x0000, 0xB0:0x0000, 0xB1:0x0000, 0xB2:0x2727, 0xB3:0x0000,
	0xC0:0x0000, 0xB8:0x0000, 0xB9:0x0000, 0xFE:0x2611, 0xFF:0x0000,
	0x88:0x2467, 0x89:0x2468, 0x3C:0x0000, 0x3B:0x0000, 0x3A:0x0000,
	0x81:0x2460, 0x86:0x2465, 0x3F:0x270D, 0x3E:0x2707, 0x3D:0x0000,
	0xE9:0x0000, 0xE8:0x2794, 0xE5:0x0000, 0xE4:0x0000, 0xE7:0x0000,
	0xE6:0x0000, 0xE1:0x0000, 0xE0:0x0000, 0xE3:0x0000, 0xE2:0x0000,
	0xF6:0x21D7, 0xEE:0x0000, 0xED:0x0000, 0xEF:0x21E6, 0xEA:0x0000,
	0xEC:0x0000, 0xEB:0x0000, 0x39:0x0000, 0x38:0x0000, 0x33:0x0000,
	0x32:0x0000, 0x31:0x0000, 0x30:0x0000, 0x37:0x2328, 0x36:0x231B,
	0x35:0x0000, 0x34:0x0000, 0xFC:0x2713, 0x60:0x264A, 0x61:0x264B,
	0x62:0x264C, 0x63:0x264D, 0x64:0x264E, 0x65:0x264F, 0x66:0x2650,
	0x67:0x2651, 0x68:0x2652, 0x69:0x2653, 0xFD:0x2612, 0x9A:0x2767,
	0x9C:0x2619, 0x9B:0x2619, 0x9E:0x2022, 0x9D:0x2767, 0x9F:0x25CF,
	0xC9:0x0000, 0xC8:0x0000, 0xC3:0x0000, 0xC2:0x0000, 0xC1:0x0000,
	0x78:0x2327, 0xC7:0x0000, 0xC6:0x0000, 0xC5:0x0000, 0xC4:0x0000,
	0xCC:0x0000, 0xCB:0x0000, 0xCA:0x0000, 0xCF:0x2766, 0xCE:0x2766,
	0xCD:0x2766, 0x99:0x2767, 0x98:0x2619, 0x91:0x277B, 0x90:0x277A,
	0x93:0x277D, 0x92:0x277C, 0x95:0x277F, 0x94:0x277E, 0x97:0x2619,
	0x96:0x2767, 0xF0:0x21E8, 0x6A:0x0026, 0x6B:0x0026, 0x6C:0x25CF,
	0x6D:0x274D, 0x6E:0x25A0, 0x6F:0x25A1, 0xF7:0x21D9, 0xF8:0x21D8,
	0xF9:0x0000, 0x8B:0x0000, 0x8C:0x2776, 0x8A:0x2469, 0xDF:0x0000,
	0x8F:0x2779, 0xDD:0x0000, 0xDE:0x0000, 0xDB:0x0000, 0xDC:0x27B2,
	0xDA:0x0000, 0x82:0x2461, 0x8D:0x2777, 0x83:0x2462, 0x8E:0x2778,
	0x80:0x24EA, 0x7F:0x0000, 0x7E:0x275E, 0x7D:0x275D, 0x7C:0x273F,
	0x7B:0x2740, 0x48:0x261F, 0x49:0x0000, 0x46:0x261E, 0x47:0x261D,
	0x44:0x0000, 0x45:0x261C, 0x42:0x0000, 0x43:0x0000, 0x40:0x270D,
	0x41:0x270C, 0xA1:0x25CB, 0xA0:0x00A0, 0xA3:0x25CB, 0xA2:0x25CB,
	0xA5:0x2609, 0x84:0x2463, 0xA7:0x25AA, 0xA6:0x274D, 0xA9:0x0000,
	0xA8:0x25A1, 0x85:0x2464, 0xAA:0x2726, 0xAC:0x2736, 0xAB:0x2605,
	0xAE:0x2738, 0xAD:0x2737, 0xAF:0x2735, 0x77:0x25C6, 0x76:0x2756,
	0x75:0x25C6, 0x74:0x25CA, 0x73:0x25CA, 0x72:0x2752, 0x71:0x2751,
	0x70:0x25A1, 0x4F:0x0000, 0x4D:0x0000, 0x4E:0x2620, 0x4B:0x263A,
	0x4C:0x2639, 0x79:0x2353, 0x4A:0x263A, 0x7A:0x2318, 0xD8:0x27A2,
	0xD9:0x0000, 0xD6:0x2326, 0xD7:0x0000, 0xD4:0x2766, 0xD5:0x232B,
	0xD2:0x2766, 0xD3:0x2766, 0xD0:0x2766, 0xD1:0x2766,
}

def wingdings_to_unicode(key):
	"""Returns the unicode value for wingding key"""
	map = WINGDINGS_MAP
	try            : value = map[key]
	except KeyError: value = map.values()[-1]
	return value

def symbol_to_unicode(key):
	"""Returns the unicode value for symbol key"""
	map = SYMBOL_MAP
	try            : value = map[key]
	except KeyError: value = map.values()[-1]
	return value


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
	
	styles    = get_styles(element)
	modifiers = get_modifiers(element)
	
	special           = intermediate.Special()
	special.modifiers = modifiers.union(styles)
	
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
		for child in element.childNodes: text = text + mask_each(child.data)
	elif element.localName == 'sym':
		symbol = element.getAttribute('w:char')
		symbol = int(symbol, 16) & 0x0FFF
		font   = element.getAttribute('w:font')
		
		if font == 'Wingdings' : to_unicode = wingdings_to_unicode
		else                   : to_unicode = symbol_to_unicode
		
		code   = to_unicode(symbol)
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
		docPr   = element.getElementsByTagName('wp:docPr').pop(0)
		alt     = docPr.getAttribute('descr');
		blip    = element.getElementsByTagName('a:blip').pop(0)
		rId     = blip.getAttribute('r:embed')
		rel     = get_relationship(rId, files)
		file    = rel.getAttribute('Target')
		imgdata = files['word/' + file]
		imgname = os.path.basename(file)
	except IndexError: return None
	
	try:#to discover dimensions
		pixels_per_emu = 12700 #Constant for word's default metric (emu)
		extent = element.getElementsByTagName('wp:extent').pop(0)
		x,y    = extent.getAttribute('cx'), extent.getAttribute('cy')
		x,y    = int(x)/pixels_per_emu, int(y)/pixels_per_emu
	except IndexError: x,y = None,None
	
	image = intermediate.Image(data=imgdata, name=imgname, height=y, width=x, descr=alt)
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

def mask_each(chars):
	"""Knocks private use unicode characters down to common usage (Microsoft
	encodes keypresses for wingding and symbol fonts in the 0xFF00-0xFFFF range)
	"""
	chars = (ord(char) for char in list(chars))
	chars = (char & 0x0FF if char > 0xF000 else char for char in chars)
	chars = (unichr(char) for char in chars)
	chars = ''.join(chars)
	return chars

def get_styles(element):
	"""
	Return style strings for a given element, elements found in run style
	elements.
	"""
	rStyles = element.getElementsByTagName('w:rStyle')
	if not rStyles: return set()
	
	styles = [style.getAttribute('w:val') for style in rStyles]
	styles = [STYLE_MAP[style] for style in styles if style in STYLE_MAP]
	return set(styles)

def get_modifiers(element):
	"""
	Return style modifiers for a given element, elements found in run property
	elements
	"""
	rPr = element.getElementsByTagName('w:rPr')
	if not rPr: return set()
	rPr = rPr.pop()
	
	modifiers = []
	for mod in rPr.childNodes:
		if mod.localName in STYLE_MAP:
			style = STYLE_MAP[mod.localName]
			#Special cases for modifiers that have attributes
			if mod.localName == 'vertAlign': 
				style = STYLE_MAP[mod.localName][mod.getAttribute('w:val')]
			modifiers.append(style)
	return set(modifiers)
