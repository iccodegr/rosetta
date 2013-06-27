import logging
from base64 import encodestring, decodestring
from tempfile import NamedTemporaryFile
from flask import Flask, request

from .converters import get_converter

log = logging.getLogger(__name__)
app = Flask("rosetta.webapi")


@app.route('/v1/conv', methods=['POST', 'GET'])
def v1_handle():
    try:
        data = request.form['data']
    except:
        return "no data was specified"
    try:
        from_type = request.form['from_type']
    except:
        return "from_type is required"
    try:
        to_type = request.form['to_type']
    except:
        return "to_type is required"
    try:
        args = request.form['args']
    except:
        args = {}

    log.debug("Converting %s => %s", from_type, to_type)

    with NamedTemporaryFile(suffix="." + from_type) as tmp:
        tmp.write(decodestring(data))
        tmp.flush()
        tmp.seek(0)
        res = get_converter(from_type, to_type)(tmp.name, to_type, args)
        if isinstance(res, basestring):
            with open(res, 'rb') as d:
                return encodestring(d.read(-1))

    return "error"


