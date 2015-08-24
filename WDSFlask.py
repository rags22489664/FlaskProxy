from __future__ import with_statement
import win32service
from subprocess import Popen, PIPE
import logging
from logging.handlers import RotatingFileHandler
from multiprocessing import Process, Lock
from multiprocessing.pool import ThreadPool

import os
import base64
import win32serviceutil
from flask import Flask, request, jsonify, send_from_directory

app = Flask(__name__, static_url_path='')
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

@app.route('/sendfile/')
def send_file():
    filename = request.args.get("filename").encode('utf8');
    macaddress = request.args.get("macaddress").encode('utf8');
    macaddress = macaddress.replace(":","")
    result = dict()
    if filename == "user-data.txt":
        if os.path.exists(os.getcwd() + "\\userdata\\" + macaddress + "\\" + filename):
            return send_from_directory(os.getcwd() + "\\userdata\\" + macaddress + "\\", filename)
        else:
            result["status"] = "Fail"
            result["status_code"] = 400
            result["message"] = "File does not exist"
            return sendresponse(result, result["status_code"])
    else:
        if filename in ["availability-zone.txt", "cloud-identifier.txt", "instance-id.txt", "local-hostname.txt", "local-ipv4.txt", "meta-data.txt", "public-hostname.txt", "public-ipv4.txt", "public-keys.txt", "service-offering.txt", "vm-id.txt"]:
            return send_from_directory(os.getcwd() + "\\metadata\\" + macaddress + "\\", filename)
        else:
            result["status"] = "Fail"
            result["status_code"] = 400
            result["message"] = "File does not exist"
            return sendresponse(result, result["status_code"])


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

@app.route("/addvmdata")
def adduserdata():

    string = request.args.get("VMData").encode('utf8');
    allEntires = string.split(";")
    result = dict()
    try:
        for entry in allEntires:
            (vmIpOrMac, folder, fileName, contents) = entry.split(',', 3)
            addVMData(vmIpOrMac, folder, fileName, contents)
    except Exception as e:
        result["status"] = "Fail"
        result["status_code"] = 400
        result["message"] = e.message
        return sendresponse(result, result["status_code"])

    result["status"] = "Pass"
    result["status_code"] = 200
    result["message"] = "Success"
    return sendresponse(result, result["status_code"])

def addVMData(vmIpOrMac, folder, fileName, contents):
    html_root = os.getcwd()
    fileName = fileName + ".txt"
    targetMetadataFile = "meta-data.txt";

    baseFolder = os.path.join(html_root, folder, vmIpOrMac)
    if not os.path.exists(baseFolder):
        os.makedirs(baseFolder)

    datafileName = os.path.join(html_root, folder, vmIpOrMac, fileName)
    metaManifest = os.path.join(html_root, folder, vmIpOrMac, targetMetadataFile)
    if folder == "userdata":
        if contents != "none":
            contents = base64.urlsafe_b64decode(contents)
        else:
            contents = ""

    try:
        f = open(datafileName, 'w')
        f.write(contents)
        f.close()
    except IOError:
        raise ErrorClass("Error while opening/writing the file " + datafileName, 400)

    if folder == "metadata" or folder == "meta-data":
        writeIfNotHere(metaManifest, fileName)

def writeIfNotHere(fileName, texts):
    if not os.path.exists(fileName):
        entries = []
    else:
        f = open(fileName, 'r')
        entries = f.readlines()
        f.close()

    texts = [ "%s\n" % t for t in texts ]
    need = False
    for t in texts:
        if not t in entries:
            entries.append(t)
            need = True

    if need:
        try:
            f = open(fileName, 'w')
            f.write(''.join(entries))
            f.close()
        except IOError:
            raise ErrorClass("Error while opening/writing the file " + fileName, 400)

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

@app.route("/ping")
def ping():
    result = dict()
    result["status"] = "OK"
    return sendresponse(result, 200)

@app.route("/deletetemplate")
def deletetemplate():
    template_uuid = request.args.get("uuid").encode('utf8');
    image_url = ""
    boot_url = ""
    client_unattended_file_url = ""
    install_unattended_file_url = ""
    image_group_name = ""
    architecture = ""
    single_image_name = ""
    if request.args.get("InstallImageFile"):
        image_url = request.args.get("InstallImageFile").encode('utf8');
    if request.args.get("InstallImageName"):
        install_image_name = request.args.get("InstallImageName").encode('utf8');
    if request.args.get("BootImageFile"):
        boot_url = request.args.get("BootImageFile").encode('utf8');
    if request.args.get("BootImageName"):
        boot_image_name = request.args.get("BootImageName").encode('utf8');
    if request.args.get("ClientUnattendFile"):
        client_unattended_file_url = request.args.get("ClientUnattendFile").encode('utf8');
    if request.args.get("ImageUnattendFile"):
        install_unattended_file_url = request.args.get("ImageUnattendFile").encode('utf8');
    if request.args.get("ImageGroupName"):
        image_group_name = request.args.get("ImageGroupName").encode('utf8');
    if request.args.get("Architecture"):
        architecture = request.args.get("Architecture").encode('utf8');
    if request.args.get("SingleImageName"):
        single_image_name = request.args.get("SingleImageName").encode('utf8');

    install_image_file_name = image_url.rpartition('\\')[2]
    boot_image_file_name = boot_url.rpartition('\\')[2]
    result = dict()
    [statusCode, out] = removeMulticastTransmission(install_image_name, image_group_name, install_image_file_name)
    if statusCode != 0:
        result["status_code"] = 400
        result["message"] = out
        result["status"] = "Fail"
        return sendresponse(result, result["status_code"])
    [statusCode, out] = removeInstallImage(install_image_name, image_group_name, install_image_file_name)
    if statusCode != 0:
        result["status_code"] = 400
        result["message"] = out
        result["status"] = "Fail"
        return sendresponse(result, result["status_code"])
    [statusCode, out] = removeBootImage(boot_image_name, architecture, boot_image_file_name)
    if statusCode != 0:
        result["status_code"] = 400
        result["message"] = out
        result["status"] = "Fail"
        return sendresponse(result, result["status_code"])
    [statusCode, out] = deleteClientUnattendedFile(client_unattended_file_url)
    if statusCode != 0:
        result["status_code"] = 400
        result["message"] = out
        result["status"] = "Fail"
        return sendresponse(result, result["status_code"])
    else:
        result["status_code"] = 200
        result["message"] = "Template Deletion Successful"
        result["status"] = "Pass"
        return sendresponse(result, result["status_code"])


def removeMulticastTransmission(install_image_name, image_group_name, install_image_file_name):

    command = "WDSUTIL /Get-MulticastTransmission /Image:\"" + install_image_name + "\" /ImageType:Install /ImageGroup:\"" + image_group_name + "\"" + " /Filename:\"" + install_image_name + ".wim\""
    proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
    out, err = proc.communicate()
    statusCode = proc.returncode
    out = filter_non_printable(out)

    if statusCode == 0:
        command = "WDSUTIL /Remove-MulticastTransmission /Image:\"" + install_image_name + "\" /ImageType:Install /ImageGroup:\"" + image_group_name + "\"" + " /Filename:\"" + install_image_name + ".wim\""
        proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
        out, err = proc.communicate()
        statusCode = proc.returncode
        out = filter_non_printable(out)

    return [statusCode, out]

def deleteClientUnattendedFile(client_unattended_file_url):

    client_unattended_file_relative_path = remoteInstallPath + "\\" + client_unattended_file_url.rpartition('\\')[2]
    command = "del /f " + client_unattended_file_relative_path
    proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
    out, err = proc.communicate()
    statusCode = proc.returncode
    out = filter_non_printable(out)

    return [statusCode, out]

def removeInstallImage(install_image_name, image_group_name, install_image_file_name):

    command = "WDSUTIL /Get-Image /Image:\"" + install_image_name + "\" /ImageType:Install /ImageGroup:\"" + image_group_name + "\"" + " /Filename:\"" + install_image_name + ".wim\""
    proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
    out, err = proc.communicate()
    statusCode = proc.returncode
    out = filter_non_printable(out)

    if statusCode == 0:
        command = "WDSUTIL /Remove-Image /Image:\"" + install_image_name + "\" /ImageType:Install /ImageGroup:\"" + image_group_name + "\"" + " /Filename:\"" + install_image_name + ".wim\""
        proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
        out, err = proc.communicate()
        statusCode = proc.returncode
        out = filter_non_printable(out)

    return [statusCode, out]

def removeBootImage(boot_image_name, architecture, boot_image_file_name):

    command = "WDSUTIL /Get-Image /Image:\"" + boot_image_name + "\" /ImageType:Boot /Architecture:" + architecture + " /Filename:\"" + boot_image_name + ".wim\""
    proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
    out, err = proc.communicate()
    statusCode = proc.returncode
    out = filter_non_printable(out)

    if statusCode == 0:
        command = "WDSUTIL /Remove-Image /Image:\"" + boot_image_name + "\" /ImageType:Boot /Architecture:" + architecture + " /Filename:\"" + boot_image_name + ".wim\""
        proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
        out, err = proc.communicate()
        statusCode = proc.returncode
        out = filter_non_printable(out)

    return [statusCode, out]


@app.route("/registertemplate")
def registertemplate():
    template_uuid = request.args.get("uuid").encode('utf8');
    if request.args.get("InstallImageFile"):
        image_url = request.args.get("InstallImageFile").encode('utf8');
    if request.args.get("InstallImageName"):
        install_image_name = request.args.get("InstallImageName").encode('utf8');
    if request.args.get("BootImageFile"):
        boot_url = request.args.get("BootImageFile").encode('utf8');
    if request.args.get("BootImageName"):
        boot_image_name = request.args.get("BootImageName").encode('utf8');
    if request.args.get("ClientUnattendFile"):
        client_unattended_file_url = request.args.get("ClientUnattendFile").encode('utf8');
    if request.args.get("ImageUnattendFile"):
        install_unattended_file_url = request.args.get("ImageUnattendFile").encode('utf8');
    if request.args.get("ImageGroupName"):
        image_group_name = request.args.get("ImageGroupName").encode('utf8');
    if request.args.get("Architecture"):
        architecture = request.args.get("Architecture").encode('utf8');
    if request.args.get("SingleImageName"):
        single_image_name = request.args.get("SingleImageName").encode('utf8');

    with lock:
        InitialTemplateDownloadRequest = template_uuid not in template_download_progress

    result = dict()

    if InitialTemplateDownloadRequest:
        p.apply_async(configureImage, args=(template_uuid, client_unattended_file_url, architecture, image_group_name, image_url, boot_url, install_unattended_file_url, single_image_name, install_image_name, boot_image_name), callback=configureImageCallBack)
        result["status"] = "InProgress"
        result["status_code"] = 200
        result["message"] = "Template registration in progress"
        with lock:
            template_download_progress[template_uuid] = result
        return sendresponse(result, 200)
    else:
        with lock:
            templateprogress = dict(template_download_progress[template_uuid])
        if "succeededOperations" in templateprogress:
            del templateprogress["succeededOperations"]
        return sendresponse(templateprogress, templateprogress["status_code"])

def configureImageCallBack(arguments):

    template_uuid = arguments["template_uuid"]
    client_unattended_file_url = arguments["client_unattended_file_url"]
    architecture = arguments["architecture"]
    image_group_name = arguments["image_group_name"]
    image_url = arguments["image_url"]
    boot_url = arguments["boot_url"]
    install_image_name = arguments["install_image_name"]
    boot_image_name = arguments["boot_image_name"]
    result = template_download_progress[template_uuid]
    if result["status_code"] == 400:
        for operation in reversed(result["succeededOperations"]):
            if operation == removeBootImage:
                boot_image_file_name = boot_url.rpartition('\\')[2]
                [statusCode, out] = removeBootImage(boot_image_name, architecture, boot_image_file_name)
            if operation == addInstallImage:
                install_image_file_name = image_url.rpartition('\\')[2]
                [statusCode, out] = removeInstallImage(install_image_name, image_group_name, install_image_file_name)
            if operation == downloadFile:
                [statusCode, out] = deleteClientUnattendedFile(client_unattended_file_url)
            if statusCode != 0:
                message = result["message"] + " and also failed while rolling back " + out
                updateTemplateDownloadProgress(template_uuid, message, "Fail", 400, result["succeededOperations"])


def configureImage(template_uuid, client_unattended_file_url, architecture, image_group_name, image_url, boot_url, install_unattended_file_url, single_image_name, install_image_name, boot_image_name):

    arguments = dict()
    arguments["template_uuid"] = template_uuid
    arguments["client_unattended_file_url"] = client_unattended_file_url
    arguments["architecture"] = architecture
    arguments["image_group_name"] = image_group_name
    arguments["image_url"] = image_url
    arguments["boot_url"] = boot_url
    arguments["install_unattended_file_url"] = install_unattended_file_url
    arguments["single_image_name"] = single_image_name
    arguments["install_image_name"] = install_image_name
    arguments["boot_image_name"] = boot_image_name
    succeededOperations = []
    [statusCode, out] = downloadFile(client_unattended_file_url, remoteInstallPath)
    if statusCode != 0:
        updateTemplateDownloadProgress(template_uuid, out, "Fail", 400, succeededOperations)
        return arguments

    succeededOperations.append(downloadFile)
    [statusCode, out] = createImageGroup(image_group_name)
    if statusCode != 0:
        updateTemplateDownloadProgress(template_uuid, out, "Fail", 400, succeededOperations)
        return arguments

    succeededOperations.append(createImageGroup)
    [statusCode, out] = addInstallImage(image_url, install_unattended_file_url, image_group_name, single_image_name, install_image_name)
    if statusCode != 0:
        updateTemplateDownloadProgress(template_uuid, out, "Fail", 400, succeededOperations)
        return arguments

    succeededOperations.append(addInstallImage)
    [statusCode, out] = addBootImage(boot_url, boot_image_name, architecture)
    if statusCode != 0:
        updateTemplateDownloadProgress(template_uuid, out, "Fail", 400, succeededOperations)
        return arguments

    succeededOperations.append(addBootImage)
    [statusCode, out] = setTransmissionTypeToImage(install_image_name, image_group_name, image_url)
    if statusCode != 0:
        updateTemplateDownloadProgress(template_uuid, out, "Fail", 400, succeededOperations)
        return arguments
    succeededOperations.append(setTransmissionTypeToImage)

    client_unattended_file_relative_path = "WdsClientUnattend" + "\\" + client_unattended_file_url.rpartition('\\')[2]
    boot_image_file_relative_path = "Boot\\" + architecture + "\\Images\\" + boot_url.rpartition('\\')[2]
    updateTemplateDownloadProgress(template_uuid, out, "Pass", 200, succeededOperations, boot_image_file_relative_path, client_unattended_file_relative_path)
    return arguments


def updateTemplateDownloadProgress(template_uuid, message, status, status_code, succeededOperations, boot_image_file_relative_path=None, client_unattended_file_relative_path=None):
    result = dict()
    result["status"] = status
    result["status_code"] = status_code
    result["succeededOperations"] = succeededOperations
    if boot_image_file_relative_path is not None:
        result["BootImagePath"] = boot_image_file_relative_path
    if client_unattended_file_relative_path is not None:
        result["WdsClientUnattend"] = client_unattended_file_relative_path
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


def setTransmissionTypeToImage(transmission_image_name, image_group_name, image_url):

    install_image_file_name = image_url.rpartition('\\')[2]
    command = "WDSUTIL /Get-MulticastTransmission /Image:\"" + transmission_image_name + "\" /ImageType:Install /ImageGroup:\"" + image_group_name + "\"" + " /Filename:\"" + transmission_image_name + ".wim\""
    proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
    out, err = proc.communicate()
    statusCode = proc.returncode
    out = filter_non_printable(out)

    if statusCode != 0:
        command = "WDSUTIL /New-MulticastTransmission /FriendlyName:\"" + transmission_image_name + " AutoCast Transmission\" /Image:\"" + transmission_image_name + "\" " \
                                                                                                                                                                     "/ImageType:Install /ImageGroup:" + image_group_name + " /TransmissionType:AutoCast "
        proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
        out, err = proc.communicate()
        statusCode = proc.returncode
        out = filter_non_printable(out)

    return [statusCode, out]

def addInstallImage(image_url, relativepath_install_unattanded_file, imagegroupname, single_image_name, install_image_name):

    install_image_file_name = image_url.rpartition('\\')[2]
    command = "WDSUTIL /Get-Image /Image:\"" + install_image_name + "\" /ImageType:Install /ImageGroup:\"" + imagegroupname + "\"" + " /Filename:\"" + install_image_name + ".wim\""

    proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
    out, err = proc.communicate()
    statusCode = proc.returncode
    out = filter_non_printable(out)

    if statusCode !=0:
        command = "WDSUTIL /Add-Image /ImageFile:\"" + image_url + "\" /ImageType:Install /UnattendFile:\"" + relativepath_install_unattanded_file + "\" /ImageGroup:" + imagegroupname
        if single_image_name:
            command = command + " /SingleImage:\"" + single_image_name + "\"" + " /Name:\"" + install_image_name + "\" /Filename:\"" + install_image_name + ".wim\""

        proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
        out, err = proc.communicate()
        statusCode = proc.returncode
        out = filter_non_printable(out)

    return [statusCode, out]

def addBootImage(boot_url, boot_image_name, architecture):

    boot_image_file_name = boot_url.rpartition('\\')[2]
    command = "WDSUTIL /Get-Image /Image:\"" + boot_image_name + "\" /ImageType:Boot /Architecture:" + architecture + " /Filename:\"" + boot_image_name + ".wim\""
    proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
    out, err = proc.communicate()
    statusCode = proc.returncode
    out = filter_non_printable(out)

    if statusCode != 0:
        command = "WDSUTIL /Add-Image /ImageFile:\"" + boot_url + "\" /ImageType:Boot" + " /Name:\"" + boot_image_name + "\"" + " /Filename:\"" + boot_image_name + ".wim\""
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
    app.run(threaded=True, host='0.0.0.0', port=8250)
