#!/usr/bin/env python3
import os
import sys
import pywinusb.hid as hid


def is_qmk_device(device):
    return device.vendor_id == 0xCB10 and device.product_id == 0x1256


def main(args: list):
    all_devices = hid.find_all_hid_devices()
    qmk_devices = list(filter(is_qmk_device, all_devices))

    pjrc_page_id = 0xFF31
    console_in_usage = (pjrc_page_id, 0x76)
    target_usage = hid.get_full_usage_id(console_in_usage[0], console_in_usage[1])
    # pyaload size is 32 bytes
    payload = [2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]

    for device in qmk_devices:
        print(device.device_path)
        try:
            device.open()
            reports = (
                device.find_input_reports()
                + device.find_output_reports()
                + device.find_feature_reports()
            )
            for report in reports:
                for k, v in report.items():
                    if v.page_id != pjrc_page_id:
                        continue
                    print("%s: %s" % (k, v))
                    if target_usage in report:
                        print("FOUND IT!!!")
                        report[target_usage] = value
                        report.send()
        finally:
            device.close()


if __name__ == "__main__":
    main(sys.argv)
