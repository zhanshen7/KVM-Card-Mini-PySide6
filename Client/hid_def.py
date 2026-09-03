import hid
from loguru import logger

product_id = 0x2107
vendor_id = 0x413D
usage_page = 0xFF00

VERBOSE = False


def set_verbose(verbose):
    global VERBOSE
    VERBOSE = verbose


h = hid.device()


# 初始化HID设备
def init_usb(vendor_id, usage_page):
    global h
    h = hid.device()
    # h.close()
    device_path = None
    for device_info in hid.enumerate():
        if (
            device_info["usage_page"] == usage_page
            and device_info["vendor_id"] == vendor_id
            and device_info["product_id"] == product_id
        ):
            device_path = device_info["path"]
    if device_path is None:
        logger.error("Device not found")
        return 1
    h.open_path(device_path)
    h.set_nonblocking(1)  # enable non-blocking mode
    return 0


def check_connection() -> bool:
    try:
        h.read(1)
        return True
    except (OSError, ValueError):
        return False


# 读写HID设备
def hid_report(buffer):
    buffer = buffer[-1:] + buffer[:-1]
    buffer[0] = 0
    if VERBOSE:
        logger.debug(f"hid < {buffer}")
    try:
        h.write(buffer)
    except (OSError, ValueError):
        logger.error("Error writing data to device")
        return 1
    return 0
