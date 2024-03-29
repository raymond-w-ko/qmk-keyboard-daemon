#!/usr/bin/env python3
import os
import sys
import time
from enum import Enum

import pywinusb.hid as hid
import win32con
import win32api
import win32gui


DEBUG = True


class APP_TYPE(Enum):
    NORMAL = 1
    GAME0 = 2
    GAME1 = 3


################################################################################

PJRC_VENDOR_ID = 0xFF31
CONSOLE_IN_USAGE = (PJRC_VENDOR_ID, 0x76)

PRIVATE_VENDOR_ID = 0xFF60
RAW_IN_USAGE = (PRIVATE_VENDOR_ID, 0x63)

################################################################################

# payload size is 32 bytes
switch_to_base_payload = [
    2,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
]
switch_to_game0_payload = [
    3,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
]
switch_to_game1_payload = [
    4,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
]

################################################################################


def is_qmk_device(device):
    return device.vendor_id == 0xCB10 and device.product_id == 0x2256


all_devices = hid.find_all_hid_devices()
qmk_devices = list(filter(is_qmk_device, all_devices))
current_app_type = None


def send_payload(payload: list):
    target_usage = hid.get_full_usage_id(RAW_IN_USAGE[0], RAW_IN_USAGE[1])
    for device in qmk_devices:
        # print(device.device_path)
        try:
            device.open()
            reports = device.find_output_reports()
            for report in reports:
                for k, v in report.items():
                    # print("%s: %s" % (k, v))
                    if target_usage in report:
                        print("found matching device, sending payload")
                        report[target_usage] = payload
                        report.send()
        finally:
            device.close()


def transition_keyboard_to_state(app_type):
    global current_app_type
    if current_app_type == app_type:
        return
    current_app_type = app_type
    if current_app_type == APP_TYPE.NORMAL:
        send_payload(switch_to_base_payload)
        print("switched to base layer")
    elif current_app_type == APP_TYPE.GAME0:
        send_payload(switch_to_game0_payload)
        print("switched to game0 layer")
    elif current_app_type == APP_TYPE.GAME1:
        send_payload(switch_to_game1_payload)
        print("switched to game1 layer")


def get_rect_size(rect):
    w = rect[2] - rect[0]
    h = rect[3] - rect[1]
    size = (w, h)
    return size


last_classname = None
last_title = None


def get_window_type(w):
    global last_classname
    global last_title
    global DEBUG

    monitor_w = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
    monitor_h = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
    rect = win32gui.GetWindowRect(w)
    size = get_rect_size(rect)
    classname = win32gui.GetClassName(w)
    title = win32gui.GetWindowText(w)
    if last_classname != classname or last_title != title:
        if DEBUG:
            print("switched to window class '%s' window title '%s'" % (classname, title, ))
        last_classname = classname
        last_title = title
    if title.endswith(" - Remote Desktop Connection"):
        return APP_TYPE.NORMAL
    elif title.endswith("TightVNC Viewer"):
        return APP_TYPE.NORMAL
    elif title.startswith("EXAPUNKS"):
        return APP_TYPE.NORMAL
    elif title.startswith("Borderlands® 3"):
        return APP_TYPE.GAME1
    elif title.startswith("SUBVERSE"):
        return APP_TYPE.GAME1
    elif title.startswith("Dyson Sphere Program"):
        return APP_TYPE.GAME1
    elif title.startswith("CODE VEIN"):
        return APP_TYPE.GAME1
    elif title.startswith("DARK SOULS"):
        return APP_TYPE.GAME0
    elif classname.startswith("POEWindowClass"):
        return APP_TYPE.GAME0
    elif "thinkorswim" in title:
        return APP_TYPE.NORMAL
    elif "Satisfactory" in title:
        return APP_TYPE.GAME1
    if size[0] == monitor_w and size[1] == monitor_h:
        return APP_TYPE.GAME0
    else:
        return APP_TYPE.NORMAL


def main(args: list):
    global DEBUG
    if "DEBUG" in args:
        DEBUG = True
    while True:
        try:
            time.sleep(0.1)
            w = win32gui.GetForegroundWindow()
            wt = get_window_type(w)
            transition_keyboard_to_state(wt)
        except KeyboardInterrupt:
            sys.exit(0)
        except:
            pass


if __name__ == "__main__":
    main(sys.argv)
