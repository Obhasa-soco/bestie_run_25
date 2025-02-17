CMD_INVENTORY = 0x01
CMD_READ_MEMORY = 0x02
CMD_WRITE_MEMORY = 0x03
CMD_READER_INFORMATION = 0x21
CMD_SET_READER_POWER = 0x2F
CMD_SET_LOCK = 0x06


class Command:
    def __init__(self, command: int, reader_address: int = 0xFF,
                 data: bytes | int | None = None):
        self.command = command
        self.reader_address = reader_address
        self.data = data
        if isinstance(data, int):
            self.data = bytearray([data])
        if data is None:
            self.data = bytearray()
        self.frame_length = 4 + len(self.data)
        self.base_data = bytearray([self.frame_length, self.reader_address, self.command])
        self.base_data.extend(self.data)

    def serialize(self) -> bytes:
        serialize = self.base_data

        # Checksum CRC-16/MCRF4XX
        value = 0xFFFF
        for d in serialize:
            value ^= d
            for _ in range(8):
                value = (value >> 1) ^ 0x8408 if value & 0x0001 else (value >> 1)
        crc_msb = value >> 0x08
        crc_lsb = value & 0xFF

        serialize = serialize + bytes([crc_lsb])
        serialize = serialize + bytes([crc_msb])
        return serialize
