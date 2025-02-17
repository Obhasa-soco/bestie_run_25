class Response:
    def __init__(self, response_bytes: bytes):
        self.response_bytes = response_bytes
        self.length = response_bytes[0]
        self.reader_address = response_bytes[1]
        self.command = response_bytes[2]
        self.status = response_bytes[3]  # Check 5. LIST OF COMMAND EXECUTION RESULT STATUS
        self.data = response_bytes[4:-2]
        self.checksum = response_bytes[-2:]

    def __str__(self) -> str:
        return_value = ''
        value = '>>> START RESPONSE ================================'
        return_value = f'{return_value}\n{value}'
        value = f'RESPONSE       >> {hex_readable(self.response_bytes)}'  # Response
        return_value = f'{return_value}\n{value}'
        value = f'READER ADDRESS >> {hex_readable(self.reader_address)}'  # Reader Address
        return_value = f'{return_value}\n{value}'
        value = f'COMMAND        >> {hex_readable(self.command)}'  # Command
        return_value = f'{return_value}\n{value}'
        value = f'STATUS         >> {hex_readable(self.status)}'  # Status
        return_value = f'{return_value}\n{value}'
        if self.data:
            value = f'DATA           >> {hex_readable(self.data)}'  # Data
            return_value = f'{return_value}\n{value}'
        value = f'CHECKSUM       >> {hex_readable(self.checksum)}'  # Checksum
        return_value = f'{return_value}\n{value}'
        value = '>>> END RESPONSE   ================================'
        return_value = f'{return_value}\n{value}'
        return return_value.strip()


def hex_readable(data: bytes | int, bytes_separator: str = " ") -> str:
    if isinstance(data, int):
        return "{:02X}".format(data)
    return bytes_separator.join("{:02X}".format(x) for x in data)