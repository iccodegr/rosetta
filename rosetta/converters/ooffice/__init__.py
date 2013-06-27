from distutils.version import LooseVersion
import os
import logging
import subprocess

from ._uno import UnoProps, OutputStream, OOFFICE, UnoUser, \
    pvalues_to_dict, UnoClient, _tcp_connection

from .formats import get_filter
from ._utils import realpath

### Now that we have found a working pyuno library, let's import some classes
# noinspection PyUnresolvedReferences
import uno, unohelper
# noinspection PyUnresolvedReferences
from com.sun.star.beans import PropertyValue
# noinspection PyUnresolvedReferences
from com.sun.star.document.UpdateDocMode import QUIET_UPDATE
# noinspection PyUnresolvedReferences
from com.sun.star.lang import DisposedException, IllegalArgumentException
# noinspection PyUnresolvedReferences
from com.sun.star.io import IOException, XOutputStream
# noinspection PyUnresolvedReferences
from com.sun.star.script import CannotConvertException
# noinspection PyUnresolvedReferences
from com.sun.star.uno import Exception as UnoException
# noinspection PyUnresolvedReferences
from com.sun.star.uno import RuntimeException
# noinspection PyUnresolvedReferences
from com.sun.star.connection import NoConnectException


class _Converter(UnoClient):
    def __init__(self, connection):
        super(_Converter, self).__init__(connection)

        self.log = logging.getLogger(_Converter.__name__)
        self.desktop = self.create_instance("com.sun.star.frame.Desktop", True)
        self.cwd = unohelper.systemPathToFileUrl(os.getcwd())


    def convert(self, inputfn, target_format, conversion_options=None):
        document = None

        self.log.debug('Input file: %s', inputfn)

        if not os.path.exists(inputfn):
            self.log.error('unoconv: file `%s\' does not exist.' % inputfn)
            return None

        try:
            ### Import phase
            phase = "import"

            ### Load inputfile
            input_md = UnoProps(Hidden=True, ReadOnly=True,
                                UpdateDocMode=QUIET_UPDATE)
            inputurl = unohelper.absolutize(self.cwd,
                                            unohelper.systemPathToFileUrl(
                                                inputfn))
            document = self.desktop.loadComponentFromURL(inputurl, "_blank", 0,
                                                         input_md)

            if not document:
                raise UnoException(
                    "The document '%s' could not be opened." % inputurl, None)

            ### Export phase
            phase = "export"

            (outputfn, ext) = os.path.splitext(inputfn)
            export_filter, out_exten = get_filter(ext, target_format)

            outputprops = UnoProps(FilterName=export_filter,
                                   OutputStream=OutputStream(),
                                   Overwrite=True)

            ### Set default filter options
            if conversion_options:
                filter_data = UnoProps(**conversion_options)
                outputprops += ( PropertyValue("FilterData",
                                               0,
                                               uno.Any(
                                                   "[]com.sun.star.beans.PropertyValue",
                                                   tuple(filter_data), ), 0), )

            outputfn = outputfn + os.extsep + out_exten

            outputurl = unohelper.absolutize(self.cwd,
                                             unohelper.systemPathToFileUrl(
                                                 outputfn))
            self.log.debug("Output file: %s" % outputfn)

            try:
                document.storeToURL(outputurl, tuple(outputprops))
            except IOException as e:
                raise UnoException(
                    "Unable to store document to %s (ErrCode %d)\n\nProperties: %s" % (
                    outputurl, e.ErrCode, outputprops), None)

            phase = "dispose"
            document.dispose()
            document.close(True)
            exitcode = 0
            return outputfn
        except SystemError as e:
            self.log.fatal(
                "unoconv: SystemError during %s phase:\n%s" % (phase, e))
            exitcode = 1

        except RuntimeException as e:
            self.log.fatal(
                "unoconv: RuntimeException during %s phase:\nOffice probably died. %s" % (
                phase, e))
            exitcode = 6

        except DisposedException as e:
            self.log.fatal(
                "unoconv: DisposedException during %s phase:\nOffice probably died. %s" % (
                phase, e))
            exitcode = 7

        except IllegalArgumentException as e:
            self.log.fatal(
                "UNO IllegalArgument during %s phase:\nSource file cannot be read. %s" % (
                phase, e))
            exitcode = 8

        except IOException as e:
        #            for attr in dir(e): print '%s: %s', (attr, getattr(e, attr))
            self.log.fatal("unoconv: IOException during %s phase:\n%s" % (
            phase, e.Message))
            exitcode = 3

        except CannotConvertException as e:
        #            for attr in dir(e): print '%s: %s', (attr, getattr(e, attr))
            self.log.fatal(
                "unoconv: CannotConvertException during %s phase:\n%s" % (
                phase, e.Message))
            exitcode = 4

        except UnoException as e:
            if hasattr(e, 'ErrCode'):
                self.log.fatal(
                    "unoconv: UnoException during %s phase in %s (ErrCode %d)" % (
                    phase, repr(e.__class__), e.ErrCode))
                exitcode = e.ErrCode
                pass
            if hasattr(e, 'Message'):
                self.log.fatal("unoconv: UnoException during %s phase:\n%s" % (
                phase, e.Message))
                exitcode = 5
            else:
                self.log.fatal("unoconv: UnoException during %s phase in %s" % (
                phase, repr(e.__class__)))
                exitcode = 2
                pass

        return exitcode


class Listener(UnoUser):
    def __init__(self, connection):
        super(Listener, self).__init__()

        self.log = logging.getLogger('ooffice.listeners')
        self.binary = OOFFICE.binary
        try:
            resolver = self.create_instance(
                "com.sun.star.bridge.UnoUrlResolver", True)
            product = self.create_instance(
                "com.sun.star.configuration.ConfigurationProvider") \
                .createInstanceWithArguments(
                "com.sun.star.configuration.ConfigurationAccess",
                UnoProps(nodepath="/org.openoffice.Setup/Product"))
            try:
                unocontext = resolver.resolve("uno:%s" % connection)
            except NoConnectException as e:
                pass
            else:
                self.log.error("Existing %s listener found, nothing to do." %
                               product.ooName)
                return

            old_params = product.ooName != "LibreOffice" or LooseVersion(
                product.ooSetupVersion) <= LooseVersion('3.3')

            params = ["-headless", "-invisible", "-nocrashreport", "-nodefault",
                      "-nologo", "-nofirststartwizard", "-norestore",
                      "-accept=%s" % connection]

            if not old_params: params = map(lambda x: "-" + x, params)
            subprocess.call([self.binary] + params, env=os.environ)

        except Exception as e:
            self.log.fatal("Launch of %s failed.\n%s" % (self.binary, e),
                           exc_info=True)
        else:
            self.log.info(
                "Existing %s listener found, nothing to do." % product.ooName)


def start_listener(port):
    return Listener(_tcp_connection('localhost', port))


def convert(fname, target, options):
    c = _Converter(_tcp_connection('localhost', 21000))
    return c.convert(fname, target, options)

