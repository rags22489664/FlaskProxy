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
thread_pool = ThreadPool(10)

REMOTE_INSTALL_PATH = "\\\\127.0.0.1\\reminst\\WdsClientUnattend"
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


def send_response(res, status_code):
    response = jsonify(res)
    response.status_code = status_code
    return response


def filter_non_printable(str):
    return ''.join([c for c in str if (31 < ord(c) < 126) or ord(c) == 9 or
                    ord(c) == 10 or ord(c) == 11 or ord(c) == 13 or
                    ord(c) == 32])


def sendfile_usage_error():
    result = dict()
    result["status"] = "Fail"
    result["status_code"] = 400
    result["message"] = "Usage: http://x.x.x.x/sendfile?" \
                        "filename=user-data.txt | availability-zone.txt | cloud-identifier.txt | instance-id.txt |" \
                        "local-hostname.txt | local-ipv4.txt | meta-data.txt | public-hostname.txt |" \
                        "public-ipv4.txt | public-keys.txt | service-offering.txt | vm-id.txt&" \
                        " macaddress=<mac address of the client machine>&"
    return send_response(result, result["status_code"])


@app.route('/sendfile/')
def send_file():
    result = dict()
    if not all([request.args.get("filename"), request.args.get("macaddress")]):
        return sendfile_usage_error()

    filename = request.args.get("filename").encode('utf8')
    macaddress = request.args.get("macaddress").encode('utf8')
    macaddress = macaddress.replace(":", "")
    if filename == "user-data.txt":
        if os.path.exists(os.getcwd() + "\\userdata\\" + macaddress + "\\" + filename):
            return send_from_directory(os.getcwd() + "\\userdata\\" + macaddress + "\\", filename)
        else:
            result["status"] = "Fail"
            result["status_code"] = 400
            result["message"] = "File does not exist"
            return send_response(result, result["status_code"])
    else:
        if filename in ["availability-zone.txt", "cloud-identifier.txt", "instance-id.txt",
                        "local-hostname.txt", "local-ipv4.txt", "meta-data.txt", "public-hostname.txt",
                        "public-ipv4.txt", "public-keys.txt", "service-offering.txt", "vm-id.txt"]:
            return send_from_directory(os.getcwd() + "\\metadata\\" + macaddress + "\\", filename)
        else:
            result["status"] = "Fail"
            result["status_code"] = 400
            result["message"] = "File does not exist"
            return send_response(result, result["status_code"])


@app.route("/wdsutil")
def wdsutil():

    command_arguments = "wdsutil"
    for param in request.query_string.split('&'):
        key = param.split('=', 1)[0]
        value = request.args.get(key)
        arg = "/" + key
        if value:
            if ' ' in value:
                arg = arg + ":\"" + value + "\""
            else:
                arg = arg + ":" + value
        command_arguments = command_arguments + " " + arg.encode('utf8')

    proc = Popen(command_arguments, shell=True, stdout=PIPE, stderr=PIPE)
    out, err = proc.communicate()
    statusCode = proc.returncode
    out = filter_non_printable(out)

    if statusCode != 0:
        raise ErrorClass(out, status_code=400)
    else:
        raise ErrorClass(out, status_code=200)


@app.route("/addvmdata")
def add_vmdata():

    string = request.args.get("VMData").encode('utf8')
    command_arguments = string.split(";")
    result = dict()
    try:
        for entry in command_arguments:
            (vm_mac_address, folder, filename, contents) = entry.split(',', 3)
            add_vmdata(vm_mac_address, folder, filename, contents)
    except Exception as e:
        result["status"] = "Fail"
        result["status_code"] = 400
        result["message"] = e.message
        return send_response(result, result["status_code"])

    result["status"] = "Pass"
    result["status_code"] = 200
    result["message"] = "Success"
    return send_response(result, result["status_code"])


def add_vmdata(vm_mac_address, folder, filename, contents):
    html_root = os.getcwd()
    filename = filename + ".txt"
    targetMetadataFile = "meta-data.txt"

    baseFolder = os.path.join(html_root, folder, vm_mac_address)
    if not os.path.exists(baseFolder):
        os.makedirs(baseFolder)

    data_filename = os.path.join(html_root, folder, vm_mac_address, filename)
    meta_manifest = os.path.join(html_root, folder, vm_mac_address, targetMetadataFile)
    if folder == "userdata":
        if contents != "none":
            contents = base64.urlsafe_b64decode(contents)
        else:
            contents = ""

    try:
        f = open(data_filename, 'w')
        f.write(contents)
        f.close()
    except IOError:
        raise ErrorClass("Error while opening/writing the file " + data_filename, 400)

    if folder == "metadata" or folder == "meta-data":
        write_if_not_here(meta_manifest, filename)


def write_if_not_here(filename, texts):
    if not os.path.exists(filename):
        entries = []
    else:
        f = open(filename, 'r')
        entries = f.readlines()
        f.close()

    texts = ["%s\n" % t for t in texts]
    need = False
    for t in texts:
        if t not in entries:
            entries.append(t)
            need = True

    if need:
        try:
            f = open(filename, 'w')
            f.write(''.join(entries))
            f.close()
        except IOError:
            raise ErrorClass("Error while opening/writing the file " + filename, 400)


@app.route("/powershell")
def powershell():

    command_arguments = "powershell Invoke-WebRequest http://10.102.153.3/cpbm/DOS.vhd -OutFile C:\\dos.vhd"
    proc = Popen(command_arguments, shell=True, stdout=PIPE, stderr=PIPE)
    out, err = proc.communicate()
    status_code = proc.returncode
    out = filter_non_printable(out)

    if status_code != 0:
        raise ErrorClass(out, status_code=400)
    else:
        raise ErrorClass(out, status_code=200)


@app.route("/ping")
def ping():
    result = dict()
    result["status"] = "OK"
    return send_response(result, 200)


def delete_template_usage_error():
    result = dict()
    result["status"] = "Fail"
    result["status_code"] = 400
    result["message"] = "Usage: http://x.x.x.x/deletetemplate?" \
                        "uuid=<template uuid>&" \
                        "ImageGroupName=<image group name>&" \
                        "ClientUnattendFile=<client unattend file path>&" \
                        "BootImageName=<unique name of the boot image>&" \
                        "InstallImageName=<unique name of the single image>&" \
                        "Architecture=<x64 or x86>"
    return send_response(result, result["status_code"])


@app.route("/deletetemplate")
def delete_template():
    if request.args.get("uuid") is None:
        return delete_template_usage_error()
    template_uuid = request.args.get("uuid").encode('utf8')
    if not all([request.args.get("InstallImageName"), request.args.get("BootImageName"),
                request.args.get("ClientUnattendFile"), request.args.get("ImageGroupName"),
                request.args.get("Architecture")]):
        return delete_template_usage_error()

    install_image_name = request.args.get("InstallImageName").encode('utf8')
    boot_image_name = request.args.get("BootImageName").encode('utf8')
    client_unattended_file_url = request.args.get("ClientUnattendFile").encode('utf8')
    image_group_name = request.args.get("ImageGroupName").encode('utf8')
    architecture = request.args.get("Architecture").encode('utf8')

    result = dict()
    [status_code, out] = remove_multicast_transmission(install_image_name, image_group_name)
    if status_code != 0:
        result["status_code"] = 400
        result["message"] = out
        result["status"] = "Fail"
        return send_response(result, result["status_code"])
    [status_code, out] = remove_install_image(install_image_name, image_group_name)
    if status_code != 0:
        result["status_code"] = 400
        result["message"] = out
        result["status"] = "Fail"
        return send_response(result, result["status_code"])
    [status_code, out] = remove_boot_image(boot_image_name, architecture)
    if status_code != 0:
        result["status_code"] = 400
        result["message"] = out
        result["status"] = "Fail"
        return send_response(result, result["status_code"])
    [status_code, out] = delete_client_unattended_file(client_unattended_file_url)
    if status_code != 0:
        result["status_code"] = 400
        result["message"] = out
        result["status"] = "Fail"
        return send_response(result, result["status_code"])
    else:
        result["status_code"] = 200
        result["message"] = "Template Deletion Successful"
        result["status"] = "Pass"
        return send_response(result, result["status_code"])


def remove_multicast_transmission(install_image_name, image_group_name):

    command = "WDSUTIL /Get-MulticastTransmission /Image:\"" + install_image_name + "\" " \
                                                 "/ImageType:Install " \
                                                 "/ImageGroup:\"" + image_group_name + "\" " + \
                                                 "/Filename:\"" + install_image_name + ".wim\""
    proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
    out, err = proc.communicate()
    status_code = proc.returncode
    out = filter_non_printable(out)

    if status_code == 0:
        command = "WDSUTIL /Remove-MulticastTransmission /Image:\"" + install_image_name + "\" " \
                                                        "/ImageType:Install " \
                                                        "/ImageGroup:\"" + image_group_name + "\" " + \
                                                        "/Filename:\"" + install_image_name + ".wim\""
        proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
        out, err = proc.communicate()
        status_code = proc.returncode
        out = filter_non_printable(out)

    return [status_code, out]


def delete_client_unattended_file(client_unattended_file_url):

    client_unattended_file_relative_path = REMOTE_INSTALL_PATH + "\\" + client_unattended_file_url.rpartition('\\')[2]
    command = "del /f " + client_unattended_file_relative_path
    proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
    out, err = proc.communicate()
    status_code = proc.returncode
    out = filter_non_printable(out)

    return [status_code, out]


def remove_install_image(install_image_name, image_group_name):

    command = "WDSUTIL /Get-Image /Image:\"" + install_image_name + "\" " \
                                 "/ImageType:Install " \
                                 "/ImageGroup:\"" + image_group_name + "\" " + \
                                 "/Filename:\"" + install_image_name + ".wim\""
    proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
    out, err = proc.communicate()
    status_code = proc.returncode
    out = filter_non_printable(out)

    if status_code == 0:
        command = "WDSUTIL /Remove-Image /Image:\"" + install_image_name + "\" " \
                                        "/ImageType:Install " \
                                        "/ImageGroup:\"" + image_group_name + "\" " + \
                                        "/Filename:\"" + install_image_name + ".wim\""
        proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
        out, err = proc.communicate()
        status_code = proc.returncode
        out = filter_non_printable(out)

    return [status_code, out]


def remove_boot_image(boot_image_name, architecture):

    command = "WDSUTIL /Get-Image /Image:\"" + boot_image_name + "\" " \
                                 "/ImageType:Boot " \
                                 "/Architecture:" + architecture + " " \
                                 "/Filename:\"" + boot_image_name + ".wim\""
    proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
    out, err = proc.communicate()
    status_code = proc.returncode
    out = filter_non_printable(out)

    if status_code == 0:
        command = "WDSUTIL /Remove-Image /Image:\"" + boot_image_name + "\" " \
                                        "/ImageType:Boot " \
                                        "/Architecture:" + architecture + " " \
                                        "/Filename:\"" + boot_image_name + ".wim\""
        proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
        out, err = proc.communicate()
        status_code = proc.returncode
        out = filter_non_printable(out)

    return [status_code, out]


def register_template_usage_error():
    result = dict()
    result["status"] = "Fail"
    result["status_code"] = 400
    result["message"] = "Usage: http://x.x.x.x/registertemplate?" \
                        "uuid=<uuid>&" \
                        "ImageGroupName=<image group name>&" \
                        "ClientUnattendFile=<client unattend file path>&" \
                        "InstallImageFile=<install image file path>&" \
                        "BootImageName=<unique name of the boot image>&" \
                        "BootImageFile=<boot image file path>&" \
                        "ImageUnattendFile=<unattended xml file path>&" \
                        "SingleImageName=<image name out of .wim file to import>&" \
                        "InstallImageName=<unique name of the single image>&" \
                        "Architecture=<x64 or x86>"
    return send_response(result, result["status_code"])


@app.route("/registertemplate")
def register_template():
    result = dict()
    if request.args.get("uuid") is None:
        return register_template_usage_error()
    template_uuid = request.args.get("uuid").encode('utf8')

    with lock:
        initial_template_download_request = template_uuid not in template_download_progress

    if initial_template_download_request:
        if not all([request.args.get("InstallImageFile"), request.args.get("InstallImageName"),
                    request.args.get("SingleImageName"), request.args.get("BootImageFile"),
                    request.args.get("BootImageName"), request.args.get("ClientUnattendFile"),
                    request.args.get("ImageUnattendFile"), request.args.get("ImageGroupName"),
                    request.args.get("Architecture")]):
            return register_template_usage_error()
        image_url = request.args.get("InstallImageFile").encode('utf8')
        install_image_name = request.args.get("InstallImageName").encode('utf8')
        boot_url = request.args.get("BootImageFile").encode('utf8')
        boot_image_name = request.args.get("BootImageName").encode('utf8')
        client_unattended_file_url = request.args.get("ClientUnattendFile").encode('utf8')
        install_unattended_file_url = request.args.get("ImageUnattendFile").encode('utf8')
        image_group_name = request.args.get("ImageGroupName").encode('utf8')
        architecture = request.args.get("Architecture").encode('utf8')
        single_image_name = request.args.get("SingleImageName").encode('utf8')

        thread_pool.apply_async(configure_image,
                      args=(template_uuid, client_unattended_file_url, architecture,
                            image_group_name, image_url, boot_url, install_unattended_file_url,
                            single_image_name, install_image_name, boot_image_name),
                      callback=configure_image_callBack)
        result["status"] = "InProgress"
        result["status_code"] = 200
        result["message"] = "Template registration in progress"
        with lock:
            template_download_progress[template_uuid] = result
        return send_response(result, 200)
    else:
        with lock:
            templateprogress = dict(template_download_progress[template_uuid])
        if "succeeded_operations" in templateprogress:
            del templateprogress["succeeded_operations"]
        return send_response(templateprogress, templateprogress["status_code"])


def configure_image_callBack(arguments):

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
        for operation in reversed(result["succeeded_operations"]):
            if operation == remove_boot_image:
                boot_image_file_name = boot_url.rpartition('\\')[2]
                [status_code, out] = remove_boot_image(boot_image_name, architecture, boot_image_file_name)
            if operation == add_install_image:
                install_image_file_name = image_url.rpartition('\\')[2]
                [status_code, out] = remove_install_image(install_image_name, image_group_name, install_image_file_name)
            if operation == download_file:
                [status_code, out] = delete_client_unattended_file(client_unattended_file_url)
            if status_code != 0:
                message = result["message"] + " and also failed while rolling back " + out
                update_template_download_progress(template_uuid, message, "Fail", 400, result["succeeded_operations"])


def configure_image(template_uuid, client_unattended_file_url, architecture, image_group_name, image_url, boot_url,
                    install_unattended_file_url, single_image_name, install_image_name, boot_image_name):

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
    succeeded_operations = []
    [status_code, out] = download_file(client_unattended_file_url, REMOTE_INSTALL_PATH)
    if status_code != 0:
        update_template_download_progress(template_uuid, out, "Fail", 400, succeeded_operations)
        return arguments

    succeeded_operations.append(download_file)
    [status_code, out] = create_image_group(image_group_name)
    if status_code != 0:
        update_template_download_progress(template_uuid, out, "Fail", 400, succeeded_operations)
        return arguments

    succeeded_operations.append(create_image_group)
    [status_code, out] = add_install_image(image_url, install_unattended_file_url,
                                          image_group_name, single_image_name, install_image_name)
    if status_code != 0:
        update_template_download_progress(template_uuid, out, "Fail", 400, succeeded_operations)
        return arguments

    succeeded_operations.append(add_install_image)
    [status_code, out] = add_boot_image(boot_url, boot_image_name, architecture)
    if status_code != 0:
        update_template_download_progress(template_uuid, out, "Fail", 400, succeeded_operations)
        return arguments

    succeeded_operations.append(add_boot_image)
    [status_code, out] = set_transmission_type_to_image(install_image_name, image_group_name, image_url)
    if status_code != 0:
        update_template_download_progress(template_uuid, out, "Fail", 400, succeeded_operations)
        return arguments
    succeeded_operations.append(set_transmission_type_to_image)

    client_unattended_file_relative_path = "WdsClientUnattend" + "\\" + client_unattended_file_url.rpartition('\\')[2]
    boot_image_file_relative_path = "Boot\\" + architecture + "\\Images\\" + boot_url.rpartition('\\')[2]
    update_template_download_progress(template_uuid, out, "Pass", 200, succeeded_operations,
                                      boot_image_file_relative_path, client_unattended_file_relative_path)
    return arguments


def update_template_download_progress(template_uuid, message, status, status_code, succeeded_operations,
                                      boot_image_file_relative_path=None, client_unattended_file_relative_path=None):
    result = dict()
    result["status"] = status
    result["status_code"] = status_code
    result["succeeded_operations"] = succeeded_operations
    if boot_image_file_relative_path is not None:
        result["BootImagePath"] = boot_image_file_relative_path
    if client_unattended_file_relative_path is not None:
        result["WdsClientUnattend"] = client_unattended_file_relative_path
    if message:
        result["message"] = message
    with lock:
        template_download_progress[template_uuid] = result


def download_file(url_to_download, path_where_to_download):

    command = "copy " + url_to_download + " " + path_where_to_download
    proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
    out, err = proc.communicate()
    status_code = proc.returncode
    out = filter_non_printable(out)

    return [status_code, out]


def set_transmission_type_to_image(transmission_image_name, image_group_name, image_url):

    command = "WDSUTIL /Get-MulticastTransmission /Image:\"" + transmission_image_name + \
              "\" /ImageType:Install /ImageGroup:\"" + image_group_name + \
              "\" /Filename:\"" + transmission_image_name + ".wim\""
    proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
    out, err = proc.communicate()
    status_code = proc.returncode
    out = filter_non_printable(out)

    if status_code != 0:
        command = "WDSUTIL /New-MulticastTransmission " \
                                "/FriendlyName:\"" + transmission_image_name + " AutoCast Transmission\" " \
                                "/Image:\"" + transmission_image_name + "\" " \
                                "/ImageType:Install /ImageGroup:" + image_group_name + " /TransmissionType:AutoCast"
        proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
        out, err = proc.communicate()
        status_code = proc.returncode
        out = filter_non_printable(out)

    return [status_code, out]


def add_install_image(image_url, relativepath_install_unattanded_file, imagegroupname, single_image_name,
                      install_image_name):

    command = "WDSUTIL /Get-Image /Image:\"" + install_image_name + "\" " \
                                 "/ImageType:Install " \
                                 "/ImageGroup:\"" + imagegroupname + "\" " + \
                                 "/Filename:\"" + install_image_name + ".wim\""

    proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
    out, err = proc.communicate()
    status_code = proc.returncode
    out = filter_non_printable(out)

    if status_code != 0:
        command = "WDSUTIL /Add-Image /ImageFile:\"" + image_url + "\" " \
                                     "/ImageType:Install " \
                                     "/UnattendFile:\"" + relativepath_install_unattanded_file + "\" " \
                                     "/ImageGroup:" + imagegroupname
        if single_image_name:
            command = command + " /SingleImage:\"" + single_image_name + "\"" + \
                                " /Name:\"" + install_image_name + "\" " \
                                "/Filename:\"" + install_image_name + ".wim\""

        proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
        out, err = proc.communicate()
        status_code = proc.returncode
        out = filter_non_printable(out)

    return [status_code, out]


def add_boot_image(boot_url, boot_image_name, architecture):

    command = "WDSUTIL /Get-Image /Image:\"" + boot_image_name + "\" " \
                                 "/ImageType:Boot " \
                                 "/Architecture:" + architecture + " " \
                                 "/Filename:\"" + boot_image_name + ".wim\""
    proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
    out, err = proc.communicate()
    status_code = proc.returncode
    out = filter_non_printable(out)

    if status_code != 0:
        command = "WDSUTIL /Add-Image /ImageFile:\"" + boot_url + "\" " \
                                     "/ImageType:Boot" + " " \
                                     "/Name:\"" + boot_image_name + "\" " + \
                                     "/Filename:\"" + boot_image_name + ".wim\""
        proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
        out, err = proc.communicate()
        status_code = proc.returncode
        out = filter_non_printable(out)

    return [status_code, out]


def create_image_group(image_group_name):

    command = "WDSUTIL /Get-ImageGroup /ImageGroup:" + image_group_name
    proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
    out, err = proc.communicate()
    status_code = proc.returncode

    if status_code != 0:
        command = "WDSUTIL /Add-ImageGroup /ImageGroup:" + image_group_name
        proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
        out, err = proc.communicate()
        status_code = proc.returncode
        out = filter_non_printable(out)

    return [status_code, out]


class PySvc(win32serviceutil.ServiceFramework):
    _svc_name_ = "CloudStack_WDS_Agent"
    _svc_display_name_ = "CloudStack WDS Agent"
    _svc_description_ = "WDS Agent for CloudStack"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)

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
    # win32serviceutil.HandleCommandLine(PySvc)
    handler = RotatingFileHandler('foo.log', maxBytes=10000, backupCount=1)
    handler.setLevel(logging.INFO)
    app.logger.addHandler(handler)
    app.run(threaded=True, host='0.0.0.0', port=8250)
