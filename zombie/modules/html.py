from output import Output
import generic

CHARSET = 'utf-8' #Character encoding for output
VERBOSE = False
ASSETS  = dict()  #Assets belonging to document (key:filename, value:data)

class Html(generic.Generic):
	def valid_out(self, output):
		""" Determines if the passed filename is accepted by this module for output """
		valid_ext = ('html', 'htm', 'xhtml', 'xml')
		output    = output.split('.')
		ext       = output.pop().lower()
		return ext in valid_ext
	
	def get_output(self, intermediate):
		""" Creates the Output object from the intermediate object
		Returns an output object, see the output module"""
		import html
		if getattr(self, 'CHARSET', False): html.CHARSET = self.CHARSET
		if getattr(self, 'VERBOSE', False): html.VERBOSE = self.VERBOSE
		
		output = get_html(intermediate)
		output = output.toprettyxml().encode(html.CHARSET, 'xmlcharrefreplace')
		output = post_processing(output)
		meta = {
			'title'   : intermediate.title,
			'author'  : intermediate.author,
			'created' : intermediate.created,
			'encoding': html.CHARSET,
			}
		output = Output(output, html.ASSETS, meta)
		return output

def post_processing(html):
	""" Processes xml and performs arbitrary post processing on it"""
	#Remove empty tags
	ignore = ('br', 'img', 'meta', 'link', 'script', 'hr', 'input', 'title', 'body', 'td')
	html   = remove_empty_tags(html, ignore)
		
	#Collapse to inline
	tags = ('a', 'strong', 'em', 'span', 'sup', 'sub')
	html = collapse_to_inline(html, tags)
	
	#Collapse to single line
	tags = ('p', 'td', 'th', 'li', 'a', 'h[\d]', 'title', 'script', 'sup', 'sub')
	html = collapse_to_single(html, tags)
	
	#remove headers from table headers
	html = clean_table(html)
	
	#fix punctuation spacing problem
	html = fix_punctuation(html)
	return html

def fix_punctuation(html):
	""" Make adjustments for punctuation where spaces may appear inappropriately"""
	import re
	search = '>[\s]+([\.\:\;\?\,\!])'
	reg    = re.compile(search)
	html   = reg.sub('>\\1', html)
	return html

def clean_table(html):
	""" Adjust a string for problems specific to tables """
	import re
	
	#Remove singular paragraphs in table cells
	search = '<(td|th)([^>]*)>[\s]*<p>(.*?)<\/p>[\s]*<\/(td|th)>'
	reg    = re.compile(search)
	result = reg.search(html)
	html   = reg.sub('<\\1\\2>\\3</\\4>', html)
	return html

def remove_empty_tags(html, ignore=()):
	""" Remove shorthand and standard empty tags.  Optional argument ignore will
	leave specified tags alone."""
	import re
	#Search for shorthand
	lkbhnd = ignore and '(?!%s)' % '|'.join(ignore) or ''
	search = '<%s[\w]+[^>]*\/>[\s]*' % lkbhnd
	reg    = re.compile(search)
	html   = reg.sub('', html)
	
	#Search for standard
	search = '<(%s[\w]+)[^>]*>[\s]*<\/\\1>[\s]*' % lkbhnd
	reg    = re.compile(search)
	html   = reg.sub('', html)
	
	return html

def collapse_to_single(html, tags):
	""" Compacts defined elements to single lines when there are no linebreaks
	within the element"""
	import re
	for tag in tags:
		search = '<(%s[\s]+[^>]*|%s)>[\s]*(.*?)[\s]*<\/(%s)>' % (tag,tag,tag)
		reg    = re.compile(search)
		html   = reg.sub('<\\1>\\2</\\3>', html)
	return html

def collapse_to_inline(html, tags):
	""" Collapses defined elements to inline formatting"""
	import re
	
	#Collapse to single line
	for tag in tags:
		search = '[\s]*<((%s)[^>]*)>[\s]*(.*?)[\s]*<\/\\2>[\s]*' % tag
		reg    = re.compile(search, re.DOTALL)
		html   = reg.sub(' <\\1>\\3</\\2> ', html)
	return html

def glue(l, o):
	""" glue the items of list l together with o. """
	from copy import copy
	o = copy(o)
	
	if l: i = l.pop(0)
	else: return []
	
	if l: return [i] + [o] + glue(l,o)
	else: return [i]

def get_dom(element, **kwargs):
	""" Given a tag, returns a new dom element for that tag """
	from xml.dom.minidom import Document
	dom = Document().createElement(element)
	for key in kwargs: dom.setAttribute(key, kwargs[key])
	return dom 

def get_text(text):
	""" Given an intermediate Text object, returns a dom text node """
	from xml.dom.minidom import Document
	if text == None: text = str()
	return Document().createTextNode(text)

def get_html(intermediate):
	""" Find the html object for a given intermediate object """
	import html
	
	ds = intermediate.__class__.__name__.lower()
	try                  : func = getattr(html, 'get_html_' + ds)
	except AttributeError: func = get_html_unknown
	return func(intermediate)

def append_to_element(element, item):
	""" Accepts a single dom item or a list of dom items and appends it to
	elements children"""
	if item == None: return 1
	if getattr(item, '__iter__', False):
		total = 0
		for i in item: total = total + append_to_element(element, i)
		return total
	element.appendChild(item)
	return 1

def get_html_document(intermediate):
	""" Return document dom node from intermediate"""
	from xml.dom.minidom import getDOMImplementation
	dimpl    = getDOMImplementation()	
	doctype  = dimpl.createDocumentType(
		'html', 
		'-//W3C//DTD XHTML 1.1//EN', 
		'http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd')	
	document = dimpl.createDocument(
		'http://www.w3.org/1999/xhtml', 
		'html', 
		doctype)
	
	html  = document.lastChild
	title = get_html_head(intermediate)
	html.appendChild(title)
	
	for element in intermediate:
		element = get_html(element)
		append_to_element(html, element)
	
	return document

def get_html_unknown(intermediate):
	""" Handle unknown objects """
	from xml.dom.minidom import Document
	name   = intermediate.__class__.__name__.lower()
	return get_html_comment('Unknown element ' + name)

def get_html_paragraph(intermediate):
	""" Return paragraph dom node from intermediate"""
	p = get_dom('p')	
	for child in intermediate:
		element = get_html(child)
		append_to_element(p, element)
	return p

def get_html_unorderedlist(intermediate):
	""" Return unordered list dom node from intermediate"""
	return get_html_list(intermediate, False)

def get_html_orderedlist(intermediate):
	""" Return ordered list dom node from intermediate"""
	return get_html_list(intermediate, True)

def get_html_table(intermediate):
	""" Return table dom node from intermediate """
	table = get_dom('table')
	for child in intermediate:
		child = get_html(child)		
		append_to_element(table, child)
	return table

def get_html_tablerow(intermediate):
	""" Return tablerow dom node from intermediate """
	tr = get_dom('tr')
	for child in intermediate:
		child = get_html(child)		
		append_to_element(tr, child)
	return tr

def get_html_tableheader(intermediate):
	""" Return tableheader dom node from intermediate """
	th = get_dom('th')
	for child in intermediate:
		child = get_html(child)		
		append_to_element(th, child)
	if intermediate.scope  : th.setAttribute('scope', intermediate.scope)
	if intermediate.colspan: th.setAttribute('colspan', intermediate.colspan)
	return th

def get_html_tablecell(intermediate):
	""" Return tablecell dom node from intermediate """
	td = get_dom('td')
	for child in intermediate:
		child = get_html(child)
		append_to_element(td, child)
	
	if intermediate.colspan: td.setAttribute('colspan', intermediate.colspan)
	if intermediate.rowspan: td.setAttribute('rowspan', intermediate.rowspan)
	return td

def get_html_image(intermediate):
	""" Return image dom node from intermediate """
	from base64    import b64encode
	from mimetypes import guess_type
	mtype, encoding = guess_type(intermediate.name)
	if mtype is None: mtype = "image/png"
	src = "data:%s;base64,%s" % (mtype, b64encode(intermediate.data))
	img = get_dom('img', src=src, alt=intermediate.descr or 'none')
	
	if intermediate.height: img.setAttribute('height', str(intermediate.height))
	if intermediate.width : img.setAttribute('width', str(intermediate.width))
	
	return img

def get_html_list(intermediate, ordered=False):
	""" Return list dom node from intermediate.  Returns ordered or unordered 
	based on keyword argument ordered=True/Fase"""
	l = get_dom(ordered and 'ol' or 'ul')
	for child in intermediate:
		child = get_html_listitem(child)		
		append_to_element(l, child)
	return l

def get_html_listitem(intermediate):
	""" Return list item dom node from intermediate """
	li = get_dom('li')
	for child in intermediate:
		child = get_html(child)		
		append_to_element(li, child)
	return li

def get_html_link(intermediate):
	""" Return anchor tag with href defined by intermediate """
	a = get_dom('a', href=intermediate.href or 'none')	
	for child in intermediate:
		child = get_html(child)		
		append_to_element(a, child)
	return a

def get_html_comment(intermediate):
	""" Create an xml dom object for a comment """
	from xml.dom.minidom import Document
	import html
	if html.VERBOSE: return Document().createComment(intermediate.data or '')
	else           : return None

def get_html_special(intermediate):
	""" Create dom object for special text objects """
	import html
	special = get_dom('span')
	for child in intermediate:
		child = get_html(child)
		append_to_element(special, child)
	
	for modifier in intermediate.modifiers:
		func = getattr(html, "get_%s" % modifier, False)
		if func: special = func(special)
		if special == None: break
		
	return special

def get_subscript(element):
	sub = get_dom('sub')
	if not element.attributes.items() and element.tagName == 'span':
		sub.childNodes = element.childNodes
	else:
		sub.appendChild(element)
	return sub

def get_superscript(element):
	sup = get_dom('sup')
	if not element.attributes.items() and element.tagName == 'span':
		sup.childNodes = element.childNodes
	else:
		sup.appendChild(element)
	return sup

def get_strong(element):
	strong = get_dom('strong')
	if not element.attributes.items() and element.tagName == 'span':
		strong.childNodes = element.childNodes
	else:
		strong.appendChild(element)
	return strong

def get_emphasis(element):
	em = get_dom('em')
	if not element.attributes.items() and element.tagName == 'span':
		em.childNodes = element.childNodes
	else:
		em.appendChild(element)
	return em

def get_underline(element):
	return get_emphasis(element)

def get_hidden(element):
	return None

def get_html_text(intermediate):
	""" Return text dom node from intermediate"""	
	data = intermediate.data.split('\n')
	data = [get_text(text) for text in data]
	br   = get_dom('br')
	data = glue(data, get_dom('br'))
	return data

def get_html_heading(intermediate):
	""" Return heading dom node from intermediate"""
	h = get_dom('h' + str(intermediate.level))
	for child in intermediate:
		child = get_html(child)		
		append_to_element(h, child)
	return h

def get_html_head(intermediate):
	""" Return head dom node from intermediate document"""
	from xml.dom.minidom import Document
	import html
	import os
	
	head = get_dom('head')
	
	if intermediate.author:
		head.appendChild(get_dom('meta', name='author', content=intermediate.author))
	
	if intermediate.created:
		head.appendChild(get_dom('meta', name='created', content=str(intermediate.created)))
	
	kwargs = {
			'http-equiv': 'Content-Type', 
			'content'   : 'text/html; charset=%s' % html.CHARSET
		}	
	head.appendChild(get_dom('meta', **kwargs))	
	head.appendChild(get_dom('meta', name='generator', content='zombie'))
	
	title = get_dom('title')
	ttext = intermediate.title or 'Untitled'
	append_to_element(title, get_text(ttext))	
	head.appendChild(title)
	
	return head

def get_html_body(intermediate):
	""" Return body dom node from intermediate"""
	body = get_dom('body')	
	for child in intermediate:		
		child = get_html(child)		
		append_to_element(body, child)	
	return body

