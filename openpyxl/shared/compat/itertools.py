try:
    from itertools import ifilter
except:
    ifilter = filter

try:
    xrange = xrange
except:
    xrange = range
