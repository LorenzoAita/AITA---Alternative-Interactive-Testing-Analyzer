import codecs
import time
import struct

from past.builtins import reduce

class Discovery(ClimaticChamber):
    run_status = False
    status_on_value = 2.3693558e-38
    status_off_value = 0.0
    sleep_read = 3


    def set_temperature(self, value):
        if value > self.configurationparameters['max temperature']:
            value = self.configurationparameters['max temperature']
        if value < self.configurationparameters['min temperature']:
            value = self.configurationparameters['min temperature']
        # address = self.connectionparameters['address']
        self.temperature_setpoint = value

        try:
            self.write(prepare_packet(504, value), force_flush=False, encode_query=False, force_unpack=False)
            return True, 1
        except Exception as e:
            return False, e.args[0]

    def get_temperature(self):
        try:
            response = self.query(prepare_packet(77), force_flush=False, encode_query=False, force_unpack=False, sleep_amount=self.sleep_read)
            self.temperature_setpoint = answer_to_float(response)
        except Exception as e:
            sleep(2)
            logger.exception(e.args[0])
        return self.temperature_setpoint

    def get_temperature_measure(self):
        try:
            response = self.query(prepare_packet(0), force_flush=False, encode_query=False, force_unpack=False, sleep_amount=self.sleep_read)
            self.temperature_measure = answer_to_float(response)
        except Exception as e:
            sleep(2)
            logger.exception(e.args[0])
        return self.temperature_measure

    def set_humidity(self, value):
        if value > 100:
            value = 100
        if value < 0:
            value = 0
        self.humidity_setpoint = value

        try:
            self.write(prepare_packet(508, value), force_flush=False, encode_query=False, force_unpack=False)
            return True, 1
        except Exception as e:
            return False, e.args[0]

    def get_humidity(self):
        try:
            response = self.query(prepare_packet(83), force_flush=False, encode_query=False, force_unpack=False, sleep_amount=self.sleep_read)
            self.humidity_setpoint = answer_to_float(response)
        except Exception as e:
            sleep(2)
            logger.exception(e.args[0])
        return self.humidity_setpoint

    def get_humidity_measure(self):
        try:
            response = self.query(prepare_packet(34), force_flush=False, encode_query=False, force_unpack=False, sleep_amount=self.sleep_read)
            self.humidity_measure = answer_to_float(response)
        except Exception as e:
            sleep(2)
            logger.exception(e.args[0])
        return self.humidity_measure

    def get_status(self):
        try:
            response = self.query(prepare_packet(73), force_flush=False, encode_query=False, force_unpack=False)
            response2 = self.query(prepare_packet(69), force_flush=False, encode_query=False, force_unpack=False)
            print(list(response))
            print(list(response2))
        except Exception as e:
            sleep(2)
            logger.exception(e.args[0])
        return response

    def set_output(self, enable=False):
        self.get_status()
        try:
            value = prepare_packet(500, self.status_on_value) if enable else prepare_packet(500, self.status_off_value)

            self.write(value, force_flush=False, encode_query=False, force_unpack=False)
            self.run_status = enable
        except Exception as e:
            return False, e.args[0]


class TH500(Discovery):

    def __init__(self, category, instrumentdescription, modelsparameters,
                 sn, code, connectionparameters, configurationparameters, caldue, atpcontext):
        ClimaticChamber.__init__(self, category, instrumentdescription, modelsparameters,
                                 sn, code, connectionparameters,
                                 configurationparameters, caldue, atpcontext)

    def connect(self):
        return ClimaticChamber.connect(self)

    def afterconnect(self, connection=None):
        pass

    def get_idn(self):
        resp = 'TH-500'
        return resp

    def set_temperature(self, value):
        if value > self.configurationparameters['max temperature']:
            value = self.configurationparameters['max temperature']
        if value < self.configurationparameters['min temperature']:
            value = self.configurationparameters['min temperature']
        value = "{:.2f}".format(value)
        self.temperature_setpoint = value

        try:
            self.write(prepare_ei_bisync_packet(value, b'S', b'L'), force_flush=False, encode_query=False, force_unpack=False)
            return True, 1
        except Exception as e:
            return False, e.args[0]

    def get_temperature_measure(self):
        try:
            response = self.query(prepare_ei_bisync_packet_read(b'P', b'V'), force_flush=False, encode_query=False, force_unpack=False)
            self.temperature_measure = ascii_to_float(response)
        except Exception as e:
            sleep(2)
            self.temperature_measure = None
            logger.exception(e.args[0])
        return self.temperature_measure

    def get_temperature(self):
        try:
            response = self.query(prepare_ei_bisync_packet_read(b'S', b'L'), force_flush=False, encode_query=False, force_unpack=False)
            self.temperature_setpoint = ascii_to_float(response)
        except Exception as e:
            sleep(2)
            # self.temperature_setpoint = None
            logger.exception(e.args[0])
        return self.temperature_setpoint

    def get_humidity(self):
        return 0.0

    def get_humidity_measure(self):
        return 0.0

    def set_output(self, enable=False):
        pass

    def get_output(self):
        pass

    def get_status(self):
        pass

    def set_humidity(self, value):
        pass


class TH1000(TH500):

    def __init__(self, category, instrumentdescription, modelsparameters,
                 sn, code, connectionparameters, configurationparameters, caldue, atpcontext):
        ClimaticChamber.__init__(self, category, instrumentdescription, modelsparameters,
                                 sn, code, connectionparameters,
                                 configurationparameters, caldue, atpcontext)

    def get_idn(self):
        resp = 'TH1000'
        return resp


# # ASCII CODES
# STX = "02"
# ETX = "03"
# EOT = "04"
# ENQ = "05"
# ACK = "06"
# NAK = "15"
#
# # Euroterm 2408 values as per documentation
# # group_id
# GID = "30"
# # unit_id
# UID = "31"


# [START] Euroterm 2408 EI-bisync protocol helpers
def prepare_ei_bisync_packet(value, first, second):
    """
    Prepare the packet to send through the EI-Bisynch protocol
    Build the packet with necessary controls character(named as in the Eurotherm 2408 manual), the user value and
    the BCC after converts the packet first in ASCII-hex and finally in bytes, ready to be sent
    :param first: the first mnemonic of the command in the euroterm 2408 manual
    :param second: the second mnemonic of the command in the euroterm 2408 manual
    :param value: user parameter float with 2 point precision
    :return:
    """
    packet = [EOT, GID, GID, UID, UID, STX, mnemonics_to_hex(first), mnemonics_to_hex(second)]
    packet.extend("{:02x}".format(ord(c)) for c in value)
    packet.append(ETX)
    BCC = calculate_bcc(bytes.fromhex(''.join(packet[6:])))
    packet.append("{:02x}".format(int(BCC) if len(BCC) > 1 else ord(BCC)))
    arr_to_str = ''.join(packet)
    hex_string = bytes.fromhex(arr_to_str)
    return hex_string


def prepare_ei_bisync_packet_read(first, second):
    query = [EOT, GID, GID, UID, UID, mnemonics_to_hex(first), mnemonics_to_hex(second), ENQ]
    query_string = ''.join(query)
    return bytes.fromhex(query_string)


def mnemonics_to_hex(value):
    return codecs.encode(value, "hex").decode('utf-8')


def ascii_to_float(response):
    """
    Decode the response and split packet received
    PV stands for Process Variable
    :param response: bytes received from the machine
    :return:
    """
    ascii_res = response.decode("ascii")
    mnemonics = ascii_res[1:3]
    res = ascii_res.split(mnemonics)[1]
    for i in range(len(res)):
        if "{:02x}".format(ord(res[i])) == ETX:
            return float(res[:i])


def calculate_bcc(values):
    return str(reduce(lambda x, y: x ^ y, values))

# [END] Euroterm 2408 EI-bisync protocols helpers


# Modbus helpers
def four_numbers(val):
    return list(reversed(list(struct.pack('f', val))))


def crc_modbus(msg):
    data = bytearray(msg)
    crc = 0xFFFF
    for pos in data:
        crc ^= pos
        for i in range(8):
            if (crc & 1) != 0:
                crc >>= 1
                crc ^= 0xA001
            else:
                crc >>= 1
    return crc & 0xFF, crc >> 8


def answer_to_float(answer):
    partial = list(answer)
    result = struct.unpack('f', bytearray([partial[6], partial[5], partial[4], partial[3]]))
    result = float(result[0])
    return round(result, 1)


def register_address(address):
    if address > 256:
        return 1, address - 256
    return 0, address


def prepare_packet(register, value=None):
    address = register_address(register)
    msg = [17, 3 if value is None else 16, address[0], address[1], 0, 2]
    if value is not None:
        msg.append(4)
        msg.extend(four_numbers(value))

    msg.extend(crc_modbus(msg))
    return msg


indexClimaticChamber = {
    "Discovery": Discovery,
    "Weiss": Weiss,
    "WeissClimeEvent": WeissClimeEvent,
    "TH500": TH500,
    "TH1000": TH1000
}