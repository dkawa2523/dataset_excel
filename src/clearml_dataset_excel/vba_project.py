from __future__ import annotations

import io

from .msovba import decompress_stream


def vba_project_has_symbol(vba_project_bin: bytes, symbol: str) -> bool:
    """
    Best-effort check if a vbaProject.bin contains a given VBA symbol.

    - Fast path: raw byte search (works for some binaries / tests)
    - Robust path: parse OLE streams and MS-OVBA decompress module source
    """
    if not symbol:
        return False

    try:
        needle_cp = symbol.encode("cp1252", errors="ignore")
    except Exception:
        needle_cp = b""
    try:
        needle_utf8 = symbol.encode("utf-8", errors="ignore")
    except Exception:
        needle_utf8 = b""

    if (needle_cp and needle_cp in vba_project_bin) or (needle_utf8 and needle_utf8 in vba_project_bin):
        return True

    try:
        import olefile  # type: ignore[import-not-found]
    except Exception:
        return False

    try:
        ole = olefile.OleFileIO(io.BytesIO(vba_project_bin))
    except Exception:
        return False

    try:
        pattern = b"\x00attribut"
        for stream_path in ole.listdir(streams=True, storages=False):
            try:
                data = ole.openstream(stream_path).read()
            except Exception:
                continue

            data_lower = data.lower()
            start = 0
            while True:
                idx = data_lower.find(pattern, start)
                if idx < 0:
                    break
                start = idx + 1
                if idx < 3:
                    continue

                compressed = data[idx - 3 :]
                try:
                    decompressed = decompress_stream(compressed)
                except Exception:
                    continue

                if needle_cp and needle_cp in decompressed:
                    return True
                if needle_utf8 and needle_utf8 in decompressed:
                    return True
        return False
    finally:
        try:
            ole.close()
        except Exception:
            pass

