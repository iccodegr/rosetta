import os, sys, glob
import logging

log = logging.getLogger("_uno_bootstrap")


def realpath(*args):
    ''' Implement a combination of os.path.join(), os.path.abspath() and
        os.path.realpath() in order to normalize path constructions '''
    ret = ''
    for arg in args:
        ret = os.path.join(ret, arg)
    return os.path.realpath(os.path.abspath(ret))


def set_office_environ(office):
    ### Set PATH so that crash_report is found
    os.environ['PATH'] = realpath(office.basepath, 'program') + os.pathsep +
                         os.environ['PATH']

    ### Set UNO_PATH so that "officehelper.bootstrap()" can find soffice executable:
    os.environ['UNO_PATH'] = office.unopath

    ### Set URE_BOOTSTRAP so that "uno.getComponentContext()" bootstraps a complete
    ### UNO environment
    if os.name in ( 'nt', 'os2' ):
        os.environ['URE_BOOTSTRAP'] = 'vnd.sun.star.pathname:' + realpath(
            office.basepath, 'program', 'fundamental.ini')
    else:
        os.environ['URE_BOOTSTRAP'] = 'vnd.sun.star.pathname:' + realpath(
            office.basepath, 'program', 'fundamentalrc')

        ### Set LD_LIBRARY_PATH so that "import pyuno" finds libpyuno.so:
        if 'LD_LIBRARY_PATH' in os.environ:
            os.environ['LD_LIBRARY_PATH'] = office.unopath + os.pathsep + \
                                            realpath(office.urepath,
                                                     'lib') + os.pathsep + \
                                            os.environ['LD_LIBRARY_PATH']
        else:
            os.environ['LD_LIBRARY_PATH'] = office.unopath + os.pathsep + \
                                            realpath(office.urepath, 'lib')

    if office.pythonhome:
        for libpath in ( realpath(office.pythonhome, 'lib'),
                         realpath(office.pythonhome, 'lib', 'lib-dynload'),
                         realpath(office.pythonhome, 'lib', 'lib-tk'),
                         realpath(office.pythonhome, 'lib', 'site-packages'),
                         office.unopath):
            sys.path.insert(0, libpath)
    else:
        ### Still needed for system python using LibreOffice UNO bindings
        ### Although we prefer to use a system UNO binding in this case
        sys.path.append(office.unopath)


def python_switch(office):
    if office.pythonhome:
        os.environ['PYTHONHOME'] = office.pythonhome
        os.environ['PYTHONPATH'] = realpath(office.pythonhome,
                                            'lib') + os.pathsep + \
                                   realpath(office.pythonhome, 'lib',
                                            'lib-dynload') + os.pathsep + \
                                   realpath(office.pythonhome, 'lib',
                                            'lib-tk') + os.pathsep + \
                                   realpath(office.pythonhome, 'lib',
                                            'site-packages') + os.pathsep + \
                                   office.unopath

    os.environ['UNO_PATH'] = office.unopath

    log.info("-> Switching from %s to %s" % (sys.executable, office.python))

    ### Set LD_LIBRARY_PATH so that "import pyuno" finds libpyuno.so:
    ld_path = os.environ.get('LD_LIBRARY_PATH', '')
    ps = os.pathsep
    uno_libs = office.unopath + ps + realpath(office.urepath, 'lib') + ps
    if not uno_libs in ld_path:
        os.environ['LD_LIBRARY_PATH'] = uno_libs + ps + ld_path
        run = 1
    else:
        run = 0

    try:
        if run:
            os.execvpe(office.python, [office.python, ] + sys.argv[0:],
                       os.environ)
    except OSError:
        ### Mac OS X versions prior to 10.6 do not support execv in
        ### a process that contains multiple threads.  Instead of
        ### re-executing in the current process, start a new one
        ### and cause the current process to exit.  This isn't
        ### ideal since the new process is detached from the parent
        ### terminal and thus cannot easily be killed with ctrl-C,
        ### but it's better than not being able to autoreload at
        ### all.
        ### Unfortunately the errno returned in this case does not
        ### appear to be consistent, so we can't easily check for
        ### this error specifically.
        ret = os.spawnvpe(os.P_WAIT, office.python,
                          [office.python, ] + sys.argv[0:], os.environ)
        sys.exit(ret)


class Office:
    def __init__(self, basepath, urepath, unopath, pyuno, binary, python,
                 pythonhome):
        self.basepath = basepath
        self.urepath = urepath
        self.unopath = unopath
        self.pyuno = pyuno
        self.binary = binary
        self.python = python
        self.pythonhome = pythonhome

    def __str__(self):
        return self.basepath

    def __repr__(self):
        return self.basepath


def find_offices():
    ret = []
    extrapaths = []

    ### Try using UNO_PATH first (in many incarnations, we'll see what sticks)
    if 'UNO_PATH' in os.environ:
        extrapaths += [os.environ['UNO_PATH'],
                       os.path.dirname(os.environ['UNO_PATH']),
                       os.path.dirname(os.path.dirname(os.environ['UNO_PATH']))]

    else:
        extrapaths += glob.glob('/usr/lib*/libreoffice*') + \
                      glob.glob('/usr/lib*/openoffice*') + \
                      glob.glob('/usr/lib*/ooo*') + \
                      glob.glob('/opt/libreoffice*') + \
                      glob.glob('/opt/openoffice*') + \
                      glob.glob('/opt/ooo*') + \
                      glob.glob('/usr/local/libreoffice*') + \
                      glob.glob('/usr/local/openoffice*') + \
                      glob.glob('/usr/local/ooo*') + \
                      glob.glob('/usr/local/lib/libreoffice*')

    ### Find a working set for python UNO bindings
    for basepath in extrapaths:
        officelibraries = ( 'pyuno.so', )
        officebinaries = ( 'soffice.bin', )
        pythonbinaries = ( 'python.bin', 'python', )
        pythonhomes = ( 'python-core-*', )

        ### Older LibreOffice/OpenOffice and Windows use basis-link/ or basis/
        libpath = 'error'
        for basis in ( 'basis-link', 'basis', '' ):
            for lib in officelibraries:
                if os.path.isfile(realpath(basepath, basis, 'program', lib)):
                    libpath = realpath(basepath, basis, 'program')
                    officelibrary = realpath(libpath, lib)
                    log.info("Found %s in %s" % (lib, libpath))
                    # Break the inner loop...
                    break
            # Continue if the inner loop wasn't broken.
            else:
                continue
                # Inner loop was broken, break the outer.
            break
        else:
            continue

        ### MacOSX have soffice binaries installed in MacOS subdirectory, not program
        unopath = 'error'
        for basis in ( 'basis-link', 'basis', '' ):
            for bin in officebinaries:
                if os.path.isfile(realpath(basepath, basis, 'program', bin)):
                    unopath = realpath(basepath, basis, 'program')
                    officebinary = realpath(unopath, bin)
                    log.info("Found %s in %s" % (bin, unopath))
                    # Break the inner loop...
                    break
            # Continue if the inner loop wasn't broken.
            else:
                continue
                # Inner loop was broken, break the outer.
            break
        else:
            continue

        ### Windows does not provide or need a URE/lib directory ?
        urepath = ''
        for basis in ( 'basis-link', 'basis', '' ):
            for ure in ( 'ure-link', 'ure', 'URE', '' ):
                if os.path.isfile(
                        realpath(basepath, basis, ure, 'lib', 'unorc')):
                    urepath = realpath(basepath, basis, ure)
                    log.info(
                        "Found %s in %s" % ('unorc', realpath(urepath, 'lib')))
                    # Break the inner loop...
                    break
            # Continue if the inner loop wasn't broken.
            else:
                continue
                # Inner loop was broken, break the outer.
            break

        pythonhome = None
        for home in pythonhomes:
            if glob.glob(realpath(libpath, home)):
                pythonhome = glob.glob(realpath(libpath, home))[0]
                log.info("Found %s in %s" % (home, pythonhome))
                break

        for pythonbinary in pythonbinaries:
            if os.path.isfile(realpath(unopath, pythonbinary)):
                log.info("Found %s in %s" % (pythonbinary, unopath))
                ret.append(Office(basepath, urepath, unopath, officelibrary,
                                  officebinary,
                                  realpath(unopath, pythonbinary), pythonhome))
        else:
            log.info("Considering %s" % basepath)
            ret.append(
                Office(basepath, urepath, unopath, officelibrary, officebinary,
                       sys.executable, None))
    return ret


OOFFICE_DISABLED = False
OOFFICE = None

#find proper openoffice installation. here we should do proper bootstrap
OOFFICE = find_offices()[0]
set_office_environ(OOFFICE)

