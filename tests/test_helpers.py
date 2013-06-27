import logging, sys

sys.path.append('.')
logging.basicConfig(level=logging.DEBUG)

from rosetta.converters import _uno

from rosetta.converters._uno import _tcp_connection
from rosetta.converters.ooffice import Listener

print sys.executable
print _uno.OFFICE, _uno.PRODUCT
Listener(_tcp_connection('localhost', 12313))

