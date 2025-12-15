from __future__ import annotations

import math
import struct


def copytoken_help(decompressed_current: int, decompressed_chunk_start: int) -> tuple[int, int, int, int]:
    """
    Compute bit masks to decode a CopyToken.

    Ported (with minor guards) from oletools.olevba.copytoken_help.
    """
    difference = decompressed_current - decompressed_chunk_start
    if difference <= 0:
        bit_count = 4
    else:
        bit_count = int(math.ceil(math.log(difference, 2)))
        bit_count = max(bit_count, 4)

    length_mask = 0xFFFF >> bit_count
    offset_mask = (~length_mask) & 0xFFFF
    maximum_length = (0xFFFF >> bit_count) + 3
    return length_mask, offset_mask, bit_count, maximum_length


def decompress_stream(compressed_container: bytes | bytearray) -> bytes:
    """
    Decompress a stream according to MS-OVBA 2.4.1.

    Ported from oletools.olevba.decompress_stream (Python implementation).
    """
    if not isinstance(compressed_container, bytearray):
        compressed_container = bytearray(compressed_container)

    if not compressed_container:
        return b""

    compressed_current = 0
    sig_byte = compressed_container[compressed_current]
    if sig_byte != 0x01:
        raise ValueError(f"invalid signature byte {sig_byte:02X}")
    compressed_current += 1

    decompressed_container = bytearray()

    while compressed_current < len(compressed_container):
        compressed_chunk_start = compressed_current
        if compressed_chunk_start + 2 > len(compressed_container):
            break

        compressed_chunk_header = struct.unpack_from("<H", compressed_container, compressed_chunk_start)[0]
        chunk_size = (compressed_chunk_header & 0x0FFF) + 3
        chunk_signature = (compressed_chunk_header >> 12) & 0x07
        if chunk_signature != 0b011:
            raise ValueError("Invalid CompressedChunkSignature in VBA compressed stream")
        chunk_flag = (compressed_chunk_header >> 15) & 0x01

        # MS-OVBA 2.4.1.3.12: max chunk size (including header) is 4098
        if chunk_flag == 1 and chunk_size > 4098:
            raise ValueError(f"CompressedChunkSize={chunk_size} > 4098 but CompressedChunkFlag == 1")
        if chunk_flag == 0 and chunk_size != 4098:
            raise ValueError(f"CompressedChunkSize={chunk_size} != 4098 but CompressedChunkFlag == 0")

        compressed_end = min(len(compressed_container), compressed_chunk_start + chunk_size)
        compressed_current = compressed_chunk_start + 2

        if chunk_flag == 0:
            decompressed_container.extend(compressed_container[compressed_current : compressed_current + 4096])
            compressed_current += 4096
            continue

        decompressed_chunk_start = len(decompressed_container)
        while compressed_current < compressed_end:
            flag_byte = compressed_container[compressed_current]
            compressed_current += 1
            for bit_index in range(8):
                if compressed_current >= compressed_end:
                    break
                flag_bit = (flag_byte >> bit_index) & 1
                if flag_bit == 0:  # LiteralToken
                    decompressed_container.append(compressed_container[compressed_current])
                    compressed_current += 1
                else:  # CopyToken
                    if compressed_current + 2 > compressed_end:
                        raise ValueError("Truncated CopyToken in VBA compressed stream")
                    copy_token = struct.unpack_from("<H", compressed_container, compressed_current)[0]
                    length_mask, offset_mask, bit_count, _ = copytoken_help(
                        len(decompressed_container), decompressed_chunk_start
                    )
                    length = (copy_token & length_mask) + 3
                    temp1 = copy_token & offset_mask
                    temp2 = 16 - bit_count
                    offset = (temp1 >> temp2) + 1
                    copy_source = len(decompressed_container) - offset
                    for index in range(copy_source, copy_source + length):
                        decompressed_container.append(decompressed_container[index])
                    compressed_current += 2

    return bytes(decompressed_container)

