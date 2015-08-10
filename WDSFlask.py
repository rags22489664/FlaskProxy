from __future__ import with_statement
import win32service
from subprocess import Popen, PIPE
import logging
from logging.handlers import RotatingFileHandler
from multiprocessing import Process, Lock
from multiprocessing.pool import ThreadPool


import win32serviceutil
from flask import Flask, request, jsonify

app = Flask(__name__)
p = ThreadPool(10)

remoteInstallPath = "\\\\127.0.0.1\\reminst\\WdsClientUnattend"
template_download_progress = dict()
lock = Lock()


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

def sendresponse(res, status_code):
    response = jsonify(res)
    response.status_code = status_code
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
        raise ErrorClass(out, status_code=400)
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
        raise ErrorClass(out, status_code=400)
    else:
        raise ErrorClass(out, status_code=200)


@app.route("/registertemplate")
def registertemplate():
    template_uuid = request.args.get("uuid").encode('utf8');
    image_url = ""
    boot_url = ""
    client_unattended_file_url = ""
    install_unattended_file_url = ""
    image_group_name = ""
    architecture = ""
    single_image_name = ""
    if request.args.get("image"):
        image_url = request.args.get("image").encode('utf8');
    if request.args.get("boot"):
        boot_url = request.args.get("boot").encode('utf8');
    if request.args.get("clientunattended"):
        client_unattended_file_url = request.args.get("clientunattended").encode('utf8');
    if request.args.get("imageunattended"):
        install_unattended_file_url = request.args.get("imageunattended").encode('utf8');
    if request.args.get("imagegroupname"):
        image_group_name = request.args.get("imagegroupname").encode('utf8');
    if request.args.get("architecture"):
        architecture = request.args.get("architecture").encode('utf8');
    if request.args.get("singleimagename"):
        single_image_name = request.args.get("singleimagename").encode('utf8');

    with lock:
        InitialTemplateDownloadRequest = template_uuid not in template_download_progress

    result = dict()

    if InitialTemplateDownloadRequest:
        p.apply_async(configureImage, args=(template_uuid, client_unattended_file_url, image_group_name, image_url, boot_url, install_unattended_file_url, single_image_name))
        result["status"] = "InProgress"
        result["status_code"] = 200
        with lock:
            template_download_progress[template_uuid] = result
        return sendresponse(result, 200)
    else:
        with lock:
            result = template_download_progress[template_uuid]
        return sendresponse(result, result["status_code"])


def configureImage(template_uuid, client_unattended_file_url, image_group_name, image_url, boot_url, install_unattended_file_url, single_image_name):

    [statusCode, out] = downloadFile(client_unattended_file_url, remoteInstallPath)
    if statusCode != 0:
        updateTemplateDownloadProgress(template_uuid, out, "Fail", 400)
        return

    [statusCode, out] = createImageGroup(image_group_name)
    if statusCode != 0:
        updateTemplateDownloadProgress(template_uuid, out, "Fail", 400)
        return

    [statusCode, out] = addImage(image_url, boot_url, install_unattended_file_url, image_group_name, single_image_name)
    if statusCode != 0:
        updateTemplateDownloadProgress(template_uuid, out, "Fail", 400)
        return

    [statusCode, out] = setTransmissionTypeToImage(single_image_name, image_group_name)
    if statusCode != 0:
        updateTemplateDownloadProgress(template_uuid, out, "Fail", 400)
        return

    updateTemplateDownloadProgress(template_uuid, out, "Pass", 200)


def updateTemplateDownloadProgress(template_uuid, message, status, status_code):
    result = dict()
    result["status"] = status
    result["status_code"] = status_code
    if message:
        result["message"] = message
    with lock:
        template_download_progress[template_uuid] = result



def downloadFile(urlToDownload, pathWhereToDownload):

    command = "copy " + urlToDownload + " " + pathWhereToDownload
    proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
    out, err = proc.communicate()
    statusCode = proc.returncode
    out = filter_non_printable(out)

    return [statusCode, out]


def setTransmissionTypeToImage(transmission_image_name, image_group_name):

    command = "WDSUTIL /New-MulticastTransmission /FriendlyName:\"" + transmission_image_name + " AutoCast Transmission\" /Image:\"" + transmission_image_name + "\" " \
                                                                                                                                                                 "/ImageType:Install /ImageGroup:" + image_group_name + " /TransmissionType:AutoCast "
    proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
    out, err = proc.communicate()
    statusCode = proc.returncode
    out = filter_non_printable(out)

    return [statusCode, out]

def addImage(image_url, boot_url, relativepath_install_unattanded_file, imagegroupname, single_image_name):

    command = "WDSUTIL /Add-Image /ImageFile:\"" + image_url + "\" /ImageType:Install /UnattendFile:\"" + relativepath_install_unattanded_file + "\" /ImageGroup:" + imagegroupname
    if single_image_name:
        command = command + " /SingleImage:\"" + single_image_name + "\""

    proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
    out, err = proc.communicate()
    statusCode = proc.returncode
    out = filter_non_printable(out)

    if statusCode != 0:
        return [statusCode, out]

    command = "WDSUTIL /Add-Image /ImageFile:\"" + boot_url + "\" /ImageType:Boot"
    proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
    out, err = proc.communicate()
    statusCode = proc.returncode
    out = filter_non_printable(out)

    return [statusCode, out]


def createImageGroup(image_group_name):

    command = "WDSUTIL /Get-ImageGroup /ImageGroup:" + image_group_name
    proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
    out, err = proc.communicate()
    statusCode = proc.returncode

    if statusCode != 0:
        command = "WDSUTIL /Add-ImageGroup /ImageGroup:" + image_group_name
        proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
        out, err = proc.communicate()
        statusCode = proc.returncode
        out = filter_non_printable(out)

    return [statusCode, out]

class PySvc(win32serviceutil.ServiceFramework):
    _svc_name_ = "CloudStack_WDS_Agent"
    _svc_display_name_ = "CloudStack WDS Agent"
    _svc_description_ = "WDS Agent for CloudStack"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self,args)

    # core logic of the service
    def SvcDoRun(self):
        self.process = Process(target=self.main)
        self.process.start()
        self.process.run()

    # called when we're being shut down
    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        self.process.terminate()
        self.ReportServiceStatus(win32service.SERVICE_STOPPED)

    def main(self):
        handler = RotatingFileHandler('foo.log', maxBytes=10000, backupCount=1)
        handler.setLevel(logging.INFO)
        app.logger.addHandler(handler)
        app.run(threaded=True)

if __name__ == '__main__':
    #win32serviceutil.HandleCommandLine(PySvc)
    handler = RotatingFileHandler('foo.log', maxBytes=10000, backupCount=1)
    handler.setLevel(logging.INFO)
    app.logger.addHandler(handler)
    app.run(threaded=True)
