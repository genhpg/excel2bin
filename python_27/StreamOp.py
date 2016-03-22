# -*- coding:utf-8 -*-

import struct

class StreamOp:
    @staticmethod
    def WriteInt(file, value):
        data = struct.pack('<i', value)
        return file.write(data)

    @staticmethod
    def ReadInt(file):
        data = struct.unpack('<i', file.read(4))
        return data

    @staticmethod
    def WriteInt8(file, value):
        data = struct.pack('<b', value)
        return file.write(data)

    @staticmethod
    def ReadInt8(file):
        data = struct.unpack('<b', file.read(1))
        return data

    @staticmethod
    def WriteInt16(file, value):
        data = struct.pack('<h', value)
        return file.write(data)

    @staticmethod
    def ReadInt16(file):
        data = struct.unpack('<h', file.read(2))
        return data

    @staticmethod
    def WriteInt64(file, value):
        data = struct.pack('<q', value)
        return file.write(data)

    @staticmethod
    def ReadInt64(file):
        data = struct.unpack('<q', file.read(8))
        return data

    @staticmethod
    def WriteFloat(file, value):
        data = struct.pack('<f', value)
        return file.write(data)

    @staticmethod
    def ReadFloat(file):
        data = struct.unpack('<f', file.read(4))
        return data

    @staticmethod
    def WriteBool(file, value):
        if value == True:
            data = struct.pack('<B', 1)
        else:
            data = struct.pack('<B', 0)
        return file.write(data)

    @staticmethod
    def ReadBool(file):
        data = struct.unpack('<B', file.read(1))
        return data

    @staticmethod
    def WriteString(file, value):
        #data = struct.pack('s', value.encode('utf-8'))
        data = value.encode('utf-8')
        StreamOp.WriteInt(file, len(data))
        return file.write(data)

    @staticmethod
    def ReadString(file):
        (len,) = StreamOp.ReadInt(file)
        bts = file.read(len)
        ret = bts.decode(encoding='utf-8')
        return ret
