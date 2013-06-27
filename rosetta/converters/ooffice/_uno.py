from _uno_bootstrap import OOFFICE

import sys
import logging
import unohelper
import uno

from com.sun.star.io import XOutputStream
from com.sun.star.beans import PropertyValue
from com.sun.star.connection import NoConnectException

log = logging.getLogger('ooffice')

### And now that we have those classes, build on them
class OutputStream(unohelper.Base, XOutputStream):
    def __init__(self):
        self.closed = 0

    def closeOutput(self):
        self.closed = 1

    def writeBytes(self, seq):
        sys.stdout.write(seq.value)

    def flush(self):
        pass


def UnoProps(**args):
    props = []
    for key in args:
        prop = PropertyValue()
        prop.Name = key
        prop.Value = args[key]
        props.append(prop)
    return tuple(props)


def pvalues_to_dict(pvals):
    d = {}
    for p in pvals:
        d[p.Name] = p.Value
    return d


class Fmt:
    def __init__(self, doctype, name, extension, summary, filter):
        self.doctype = doctype
        self.name = name
        self.extension = extension
        self.summary = summary
        self.filter = filter

    def __str__(self):
        return "%s [.%s]" % (self.summary, self.extension)

    def __repr__(self):
        return "%s/%s" % (self.name, self.doctype)


def create_instance(class_, context=True):
    ctx = uno.getComponentContext()
    if not context:
        return ctx.ServiceManager.createInstance(class_)
    else:
        return ctx.ServiceManager.createInstanceWithContext(class_, ctx)


class UnoUser(object):
    def __init__(self):
        self.context = uno.getComponentContext()
        self.srvmgr = self.context.ServiceManager

    def create_instance(self, class_, with_context=False):
        if not with_context:
            return self.srvmgr.createInstance(class_)
        else:
            return self.srvmgr.createInstanceWithContext(class_, self.context)


class UnoClient(UnoUser):
    def __init__(self, connection):
        super(UnoClient, self).__init__()
        self.log = logging.getLogger('UnoClient[%s]' % connection)
        self.connection = connection
        self.__connect()

    def __connect(self):
        resolver = self.create_instance("com.sun.star.bridge.UnoUrlResolver",
                                        True)
        self.log.debug('Connection type: %s' % self.connection)
        try:
            self.context = resolver.resolve("uno:%s" % self.connection)
            self.srvmgr = self.context.ServiceManager
        except NoConnectException as e:
            self.log.error("Existing listener not found.", exc_info=True)
        if not self.context:
            self.log.fatal("Unable to connect or start own listener. Aborting.")


OFFICE = OOFFICE
__CTX = uno.getComponentContext()
__SRVMGR = __CTX.ServiceManager
PRODUCT = create_instance(
    "com.sun.star.configuration.ConfigurationProvider").createInstanceWithArguments(
    "com.sun.star.configuration.ConfigurationAccess",
    UnoProps(nodepath="/org.openoffice.Setup/Product"))


def _tcp_connection(hostname, port):
    return "socket,host=%s,port=%s;urp;StarOffice.ComponentContext" % (
    hostname, port)
