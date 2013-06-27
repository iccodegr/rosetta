import os


def realpath(*args):
    ''' Implement a combination of os.path.join(), os.path.abspath() and
        os.path.realpath() in order to normalize path constructions '''
    ret = ''
    for arg in args:
        ret = os.path.join(ret, arg)
    return os.path.realpath(os.path.abspath(ret))
