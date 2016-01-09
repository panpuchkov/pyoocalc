# -*- coding: utf-8 -*-

import os
import subprocess
import sys
import time
import uno

sys.path.append('./../')
import pyloo

NoConnectionException = uno.getClass(
        "com.sun.star.connection.NoConnectException")

###############################################################################


def start_office(timeout=30, attempt_period=0.1,
                 office='soffice \
--accept="socket,host=localhost,port=2002;urp;"'):
    """
    Starts Libre/Open Office with a listening socket.

    @type  timeout: int
    @param timeout: Timeout for starting Libre/Open Office in seconds

    @type  attempt_period: int
    @param attempt_period: Timeout between attempts in seconds

    @type  office: string
    @param office: Libre/Open Office startup string
    """
    ###########################################################################
    def start_office_instance(office):
        """
        Starts Libre/Open Office with a listening socket.

        @type  office: string
        @param office: Libre/Open Office startup string
        """
        # Fork to execute Office
        if os.fork():
            return

        # Start OpenOffice.org and report any errors that occur.
        try:
            retcode = subprocess.call(office, shell=True)
            if retcode < 0:
                print (sys.stderr,
                       "Office was terminated by signal",
                       -retcode)
            elif retcode > 0:
                print (sys.stderr,
                       "Office returned",
                       retcode)
        except OSError as e:
            print (sys.stderr, "Execution failed:", e)

        # Terminate this process when Office has closed.
        raise SystemExit()

    ###########################################################################
    waiting = False
    doc = None
    try:
        doc = pyloo.Document()
    except NoConnectionException as e:
        waiting = True
        start_office_instance(office)

    if waiting:
        steps = int(timeout/attempt_period)
        for i in range(steps + 1):
            try:
                doc = pyloo.Document()
                break
            except NoConnectionException as e:
                time.sleep(attempt_period)
    del doc

###############################################################################

start_office()
print ("Office started")
try:
    doc = pyloo.Document()
    file_name = os.getcwd() + "/example.ods"
    doc.open_document(file_name)
except NoConnectionException as e:
    print (e)
