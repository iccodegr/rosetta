import logging, sys

sys.path.append('.')
logging.basicConfig(level=logging.DEBUG)

from rosetta.converters.ooffice._uno import _tcp_connection, create_instance, UnoClient, pvalues_to_dict
from rosetta.converters.ooffice import Listener, _Converter


print sys.executable
c = _Converter(_tcp_connection('localhost', 12313))
c.convert('simple.docx', 'pdf', {
    'Watermark': 'test'
})

c.convert('test.txt', 'pdf', {
    'Watermark': 'test'
})
#cl = UnoClient(_tcp_connection('localhost',12313))
#s = cl.create_instance("com.sun.star.document.FilterFactory", True)
#ss =  s.getElementNames()
#import pprint
#with open('f.txt', 'w') as f:
#    for n in ss:
#        pprint.pprint(pvalues_to_dict(s.getByName(n)), stream=f)
