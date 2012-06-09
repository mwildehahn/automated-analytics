import numpy as np
import getopt

# These are really outdated and written while I was still learning python,
# attempting to merge over matlab functions. `accumarray` is the only one that
# is still used heavily.

def ismember(list1,list2,just_tf=0):
	aL1 = np.array(list1)
	aL2 = np.array(list2)
	temp_tf = [i in aL2 for i in aL1]
	tf = np.array(temp_tf)
	loc = np.array(range(len(aL1)))[tf]
	if just_tf:
		return tf
	else:
		return tf,loc

def logical_tf(type,*criteria):
    temp = criteria[0]
    if type == 'and':
      for criterion in criteria:
        for index, i in enumerate(criterion):
          if temp[index] != i:
            temp[index] = False
    elif type == 'or':
      for criterion in criteria:
        for index, i in enumerate(criterion):
          if i:
            temp[index] = i
    else:
      raise TypeError("Type must be 'and' or 'or'")
    return temp

def accumarray(accum_by, accum_value, function=sum, none=0):
    # hack for count function
    if function == '@count':
        function = len
	# check to make sure all the inputs are the same length
    same_length = [1 for i in map(len, accum_by)
        if i == len(accum_value)]
    if sum(same_length) == len(accum_by):
        # account for python indices starting at 0
        altered_accum_by = map(lambda x: x+1, accum_by)
        # create empty matrix of shape accum_by[0] x accum_by[1]
        out_array = np.empty((map(max,altered_accum_by)),'object')
        for i in range(len(accum_by[0])):
            index = [0]*len(accum_by)
            for j in range(len(accum_by)):
                index[j] = accum_by[j][i]
            index = tuple(index)
            if isinstance(out_array[index], list):
                out_array[index].append(accum_value[i])
            else:
                out_array[index] = []
                out_array[index].append(accum_value[i])
        for i in range(len(accum_by[0])):
            index = [0]*len(accum_by)
            for j in range(len(accum_by)):
                index[j] = accum_by[j][i]
            index = tuple(index)
            if isinstance(out_array[index], list):
                out_array[index] = function(out_array[index])
            else:
                pass
        if none or len(np.shape(out_array)) == 1:
            return out_array
        else:
            for row in out_array:
                tf = logical_tf('and', row < 0)
                row[tf] = 0
            return out_array
    else:
        print 'Check the inputs, they weren\'t the same length'

def sort_table(unsortedtable,sortcolumn=0,type='asc'):
	if '|S' in str(unsortedtable.dtype):
		float_vec = np.vectorize(float)
		if type == 'asc':
			sortedtable = unsortedtable[float_vec(unsortedtable[:,sortcolumn]).argsort(),:]
		elif type == 'desc':
			sort_index = float_vec(unsortedtable[:,sortcolumn]).argsort().tolist()
			sort_index.reverse()
			sortedtable = unsortedtable[sort_index,:]
	else:
		if type == 'asc':
			sortedtable = unsortedtable[unsortedtable[:,sortcolumn].argsort(),:]
		elif type == 'desc':
			sort_index = unsortedtable[:,sortcolumn].argsort().tolist()
			sort_index.reverse()
			sortedtable = unsortedtable[sort_index,:]
	return sortedtable

def rowcol(input):
	x,y = np.shape(input)
	[r,c] = np.meshgrid(range(y),range(x))
	return r,c

def stack_dict(dictionary,headers = 0):
	if headers != 0:
		header_row = dictionary.keys()
		table = np.column_stack((dictionary.values()))
		table = np.vstack((header_row,table))
	else:
		table = np.column_stack((dictionary.values()))
	return table
