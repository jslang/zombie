class Generic:
	""" Generic container for document tree """
	def __init__(self):
		self.children  = list()
		self.modifiers = set()
	
	def append(self, item):
		if isinstance(item, list) or isinstance(item, tuple):
			for i in item: Generic.append(self, i)
		elif item == None           : return
		else                        : self.children.append(item)
		
	def __repr__(self):
		return self.get_tree()
	
	def get_tree(self, lvl=0):
		indent = '\t' * lvl
		out    = indent + self.__class__.__name__ 
		out    += self.children and ' (\n' or '\n'
		for child in self.children:	out += child.get_tree(lvl+1)
		if self.children: out += indent + ')\n'
		return out
	
	def __iter__(self):
		return self.children.__iter__()
	
	def __len__(self):
		return len(self.children)
	
	def __getitem__(self, key):
		return self.children[key]
	
	def __setitem__(self, key, val):
		self.children[key] = val
	
	def __contains__(self, item):
		return item in self.children

class Document(Generic):
	""" Root Element in a document tree, contains meta information for document """
	def __init__(self, title=None, author=None, created=None):
		Generic.__init__(self)
		self.title   = title
		self.author  = author
		self.created = created

class Body(Generic):
	""" Body container """
	def __init__(self):
		Generic.__init__(self)

class Comment(Generic):
	""" Comments """
	def __init__(self, text=None):
		Generic.__init__(self)
		self.data = text or ''

class Text(Generic):
	""" Text elements in doc tree """
	def __init__(self, text=None):
		Generic.__init__(self)
		self.data = text or ''
	
	def __add__(self, text):
		if isinstance(text, Text): text = self.data + text.data
		else                     : text = self.data + text
		return Text(text)
	
	def __repr__(self):
		return self.data

class Special(Generic):
	""" Text that is different from its neighbors """
	def __init__(self): Generic.__init__(self)
	
	def __add__(self, x):		
		if not isinstance(x, Special): raise TypeError('Expected argument of type Special')
		from modules import compact
		self.children  = self.children + x.children
		self.children  = compact(self.children, lambda x,y: isinstance(x, Text) and isinstance(y, Text))
		self.modifiers = list(set(x.modifiers).union(self.modifiers))
		return self

class Heading(Generic):
	""" Heading elements in doc tree """
	def __init__(self, level=0):
		Generic.__init__(self)
		self.level = level

class Paragraph(Generic):
	""" Paragraph elements in doc tree """
	def __init__(self, text=None):
		Generic.__init__(self)
		if text: self.append(text)
	
	def append(self, item):
		if len(self.children) and isinstance(item, Text) and isinstance(self.children[-1], Text):
			self.children[-1] = self.children[-1] + item
		else: Generic.append(self, item)

class List(Generic):
	""" Generic List Element """
	def __init__(self): Generic.__init__(self)
	
	def __add__(self, x):
		if isinstance(x, List):	self.append(x.children)
		return self

class ListItem(Generic):
	""" Generic List Item Element """
	def __init__(self): Generic.__init__(self)

class OrderedList(List):
	""" Numbered Lists """
	def __init__(self):	Generic.__init__(self)

class UnorderedList(List):
	""" Bulleted Lists """
	def __init__(self):	Generic.__init__(self)

class Table(Generic):
	""" Table data """
	def __init__(self, caption=None):
		Generic.__init__(self)
		self.caption = caption

class TableRow(Generic):
	""" Table row """
	def __init__(self):
		Generic.__init__(self)

class TableCell(Generic):
	""" Table cells """
	def __init__(self, colspan=None, rowspan=None):
		Generic.__init__(self)
		self.colspan = colspan
		self.rowspan = rowspan

class TableHeader(TableCell):
	""" Table headers """
	def __init__(self, colspan=None, rowspan=None, scope='none'): 
		TableCell.__init__(self, colspan, rowspan)
		self.scope = scope
	
class Media(Generic):
	""" Generic Media element """
	def __init__(self, data=None, name='', descr=''):
		Generic.__init__(self)
		self.data  = data
		self.name  = name
		self.descr = descr

class Image(Media):
	""" Images linked in a document """
	def __init__(self, data=None, name='', descr='', height=None, width=None):
		Media.__init__(self, data, name, descr)
		self.height = height
		self.width  = width

class Link(Generic):
	""" Links to other resources or specific locations in a document """
	def __init__(self, href=None):
		Generic.__init__(self)
		self.href = href
