import logging
from logging.handlers import RotatingFileHandler
from subprocess import Popen, PIPE
from flask import Flask, request, jsonify
import re
import sys

app = Flask(__name__)


class ErrorClass(Exception):
    status_code = 400

    def __init__(self, message, status_code=None, payload=None):
        Exception.__init__(self)
        self.message = message
        if status_code is not None:
            self.status_code = status_code
        self.payload = payload

    def to_dict(self):
        rv = dict(self.payload or ())
        print "message", self.message
        rv['message'] = self.message
        return rv


@app.errorhandler(ErrorClass)
def handle_invalid_usage(error):
    response = jsonify(error.to_dict())
    response.status_code = error.status_code
    return response


def filter_non_printable(str):
    return ''.join([c for c in str if (31 < ord(c) < 126) or ord(c) == 9 or ord(c) == 10 or ord(c) == 11 or ord(c) == 13 or ord(c) == 32])


@app.route("/wdsutil")
def wdsutil():

    commandArgs = "wdsutil"
    for param in request.query_string.split('&'):
        key = param.split('=', 1)[0]
        value = request.args.get(key);
        arg = "/" + key;
        if value:
            if ' ' in value:
                arg = arg + ":\"" + value + "\""
            else:
                arg = arg + ":" + value
        commandArgs = commandArgs + " " + arg.encode('utf8');

    proc = Popen(commandArgs, shell=True, stdout=PIPE, stderr=PIPE)
    out, err = proc.communicate()
    statusCode = proc.returncode
    out = filter_non_printable(out);

    if statusCode != 0:
        raise ErrorClass(out, status_code=410)
    else:
        raise ErrorClass(out, status_code=200)

@app.route("/powershell")
def powershell():

    commandArgs = "powershell Invoke-WebRequest http://10.102.153.3/cpbm/DOS.vhd -OutFile C:\\dos.vhd"
    proc = Popen(commandArgs, shell=True, stdout=PIPE, stderr=PIPE)
    out, err = proc.communicate()
    statusCode = proc.returncode
    out = filter_non_printable(out);

    if statusCode != 0:
        raise ErrorClass(out, status_code=410)
    else:
        raise ErrorClass(out, status_code=200)


if __name__ == "__main__":
    handler = RotatingFileHandler('foo.log', maxBytes=10000, backupCount=1)
    handler.setLevel(logging.INFO)
    app.logger.addHandler(handler)
    app.run()
