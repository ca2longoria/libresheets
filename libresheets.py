
import re
import copy

from zipfile import ZipFile
import xml.etree.ElementTree as ET


class InvalidSpreadsheetFileException(Exception): pass

def _eldig(n,check):
	''' Recursively filter xml.etree.ElementTree nodes for matches of check(). '''
	r = []
	def rec(n,path=None,depth=0):
		path = path and path or []
		try:
			if check(n,path,depth):
				r.append(n)
		except TypeError as e:
			if 'positional argument but' in str(e):
				if check(n):
					r.append(n)
			else:
				raise e
		path.append(n)
		for c in n:
			rec(c,path,depth+1)
		path.pop()
	rec(n)
	return r

class SimpleSheets:
	'''
	Open LibreOffice ODS spreadsheet file, and parse simple table text data.
	'''
	def __init__(self,target):
		self._target = target
		self._data = None
	
	def sheets(self):
		''' Return and cache a dict of a table per sheet with basic text data. '''
		if not self._data:
			self._data = self._parse_data()
		return copy.deepcopy(self._data)
	
	def clean_sheets(self):
		''' Return self.sheets() but with tuple keys replaced with strings. '''
		t = self.sheets()
		ret = {}
		for k in t:
			d = {}
			for k2 in t[k]:
				key = k2
				if type(k2) != str:
					key = '%i,%i' % k2
				d[key] = t[k][k2]
			ret[k] = d
		return ret
	
	def _parse_data(self):
		''' Get the dict of tables. '''
		out = {}
		s = self._zip_data(self._target)
		root = ET.fromstring(s)
		x = 1
		for t in _eldig(root,lambda n:re.search(r'table$',n.tag)):
			key = 'sheet%03s' % (x,)
			for k,v in t.attrib.items():
				if re.search(r'name$',k):
					key = v
					break
			out[key] = dict( ( ((x,y),s) for y,x,s in self._el_cells(t) ) )
			x += 1
		return out
	
	def _el_cells(self,root):
		''' Yield all cells of rows in the sheet. '''
		rows = list(_eldig(root,lambda n:re.search(r'table-row$',n.tag)))
		row = 0
		for n in rows:
			col = 0
			for c in _eldig(n,lambda c:re.search(r'table-cell$',c.tag)):
				s = self._cell_text(c)
				if s:
					yield (row,col,s)
				col += 1
			row += 1
	
	def _cell_text(self,cell):
		''' Get text within cell (join <text/> nodes). '''
		sr = []
		for c in _eldig(cell,lambda n:'text' in n.tag):
			sr.append(c.text)
		return '\n'.join(sr)
	
	def _zip_data(self,target=None):
		''' Get text of content.xml within the target .ods (zip) file. '''
		target = target and target or self._target
		d = None
		with ZipFile(target,'r') as f:
			r = f.namelist()
			if not 'content.xml' in r:
				raise InvalidSpreadsheetFileException(self._target)
			d = f.open('content.xml').read()
		return d


if __name__ == '__main__':
	
	import sys
	import json
	
	target = sys.argv[1]
	f = SimpleSheets(target)
	print(json.dumps(f.clean_sheets(),indent=2))
	
