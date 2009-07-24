__all__ = ['docx', 'html']

def compact(series, test):
	""" Given a series, will find elements in that series that are the same type 
	and neighbors, and will merge them via the addition operator (type must 
	override addition method).  Returns this compacted series."""
	new_series = list()
	for item in series:
		if new_series and test(item, new_series[-1]):
			new_series[-1] = new_series[-1] + item
		else:
			new_series.append(item)
	return new_series