import sys
from os.path import dirname, join

__LIB_DIR__ = dirname(__file__)
sys.path.append(__LIB_DIR__)
sys.path.append(join(__LIB_DIR__, 'flask'))
sys.path.append(join(__LIB_DIR__, 'werkzeug'))
sys.path.append(join(__LIB_DIR__, 'itsdangerous'))

import flask
