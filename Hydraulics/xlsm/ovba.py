"""Minimal writer for an Excel ``vbaProject.bin`` (source-only VBA project, MS-OVBA / MS-CFB).

    make_vba_project(modules)  ->  bytes
        modules = [(name, kind, source_text)]   kind: 'workbook' | 'sheet' | 'std'

No p-code is written: Excel recompiles the module source when the file is opened.
"""
import struct

# ── MS-OVBA compression ────────────────────────────────────────────────────
def _compress_chunk(chunk: bytes) -> bytes:
    out = bytearray()
    pos = 0
    n = len(chunk)
    while pos < n:
        flag_pos = len(out)
        out.append(0)
        flags = 0
        for bit in range(8):
            if pos >= n:
                break
            diff = pos
            best_len, best_off = 0, 0
            if diff > 0:
                bit_count = max((diff - 1).bit_length(), 4)
                max_len = (0xFFFF >> bit_count) + 3
                start = 0
                # greedy longest match (chunk is <= 4096 bytes, brute force is fine)
                b0 = chunk[pos]
                i = chunk.find(bytes([b0]), 0, pos)
                while i != -1:
                    ln = 0
                    while pos + ln < n and ln < max_len and chunk[i + ln] == chunk[pos + ln]:
                        ln += 1
                    if ln > best_len:
                        best_len, best_off = ln, pos - i
                    i = chunk.find(bytes([b0]), i + 1, pos)
            if best_len >= 3:
                bit_count = max((diff - 1).bit_length(), 4)
                tok = ((best_off - 1) << (16 - bit_count)) | (best_len - 3)
                out += struct.pack("<H", tok)
                flags |= 1 << bit
                pos += best_len
            else:
                out.append(chunk[pos])
                pos += 1
        out[flag_pos] = flags
    return bytes(out)


def compress(data: bytes) -> bytes:
    out = bytearray([1])
    for i in range(0, max(len(data), 1), 4096):
        chunk = data[i:i + 4096]
        comp = _compress_chunk(chunk)
        if len(comp) + 2 > 4098 or len(chunk) < 1:
            # fall back to a raw chunk (needs a full 4096 bytes)
            raw = chunk + b"\x00" * (4096 - len(chunk))
            out += struct.pack("<H", (4096 + 2 - 3) | (0b011 << 12)) + raw
        else:
            out += struct.pack("<H", (len(comp) + 2 - 3) | (0b011 << 12) | 0x8000) + comp
    return bytes(out)


# ── project streams ────────────────────────────────────────────────────────
_BASE = {"workbook": "0{00020819-0000-0000-C000-000000000046}",
         "sheet": "0{00020820-0000-0000-C000-000000000046}"}
_ID = "{8D807122-0657-42C8-BC6F-4B5FD08031C9}"            # CMG/DPB/GC below are keyed to this ID
_CMG, _DPB, _GC = "DBD966ECE4F0E4F0E4F0E4F0", "5557E86E18E919E919E9", "CFCD72F0961011111111EE"

_REFS = (  # stdole + Office references (verbatim from an Excel-saved project)
    b"\x16\x00\x06\x00\x00\x00stdole>\x00\x0c\x00\x00\x00s\x00t\x00d\x00o\x00l\x00e\x00"
    b"\x0d\x00h\x00\x00\x00^\x00\x00\x00*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#"
    b"C:\\WINDOWS\\system32\\stdole2.tlb#OLE Automation\x00\x00\x00\x00\x00\x00"
    b"\x16\x00\x06\x00\x00\x00Office>\x00\x0c\x00\x00\x00O\x00f\x00f\x00i\x00c\x00e\x00"
    b"\x0d\x00\x9e\x00\x00\x00\x94\x00\x00\x00*\\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.0#0#"
    b"C:\\Program Files\\Common Files\\Microsoft Shared\\OFFICE12\\MSO.DLL#Microsoft Office 12.0 Object Library"
    b"\x00\x00\x00\x00\x00\x00")


def _rec(rid, payload=b""):
    return struct.pack("<HI", rid, len(payload)) + payload


def _dir_stream(mods):
    d = bytearray()
    d += _rec(0x01, struct.pack("<I", 1))                       # SYSKIND win32
    d += _rec(0x02, struct.pack("<I", 0x409))
    d += _rec(0x14, struct.pack("<I", 0x409))
    d += _rec(0x03, struct.pack("<H", 1252))
    d += _rec(0x04, b"VBAProject")
    d += _rec(0x05) + _rec(0x40)
    d += _rec(0x06) + _rec(0x3D)
    d += _rec(0x07, struct.pack("<I", 0))
    d += _rec(0x08, struct.pack("<I", 0))
    d += struct.pack("<HIIH", 0x09, 4, 0x52671F51, 0x0020)        # VERSION
    d += _rec(0x0C) + _rec(0x3C)
    d += _REFS
    d += _rec(0x0F, struct.pack("<H", len(mods)))
    d += _rec(0x13, struct.pack("<H", 0xFFFF))
    for name, kind, _ in mods:
        nm = name.encode("ascii")
        d += _rec(0x19, nm) + _rec(0x47, name.encode("utf-16-le"))
        d += _rec(0x1A, nm) + _rec(0x32, name.encode("utf-16-le"))
        d += _rec(0x1C) + _rec(0x48)
        d += _rec(0x31, struct.pack("<I", 0))                   # no p-code: source starts at 0
        d += _rec(0x1E, struct.pack("<I", 0))
        d += _rec(0x2C, struct.pack("<H", 0xFFFF))
        d += _rec(0x21 if kind == "std" else 0x22)
        d += _rec(0x2B)
    d += _rec(0x10)
    return compress(bytes(d))


def _module_source(name, kind, text):
    head = f'Attribute VB_Name = "{name}"\r\n'
    if kind in _BASE:
        head += (f'Attribute VB_Base = "{_BASE[kind]}"\r\nAttribute VB_GlobalNameSpace = False\r\n'
                 'Attribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = True\r\n'
                 'Attribute VB_Exposed = True\r\nAttribute VB_TemplateDerived = False\r\n'
                 'Attribute VB_Customizable = True\r\n')
    body = text.replace("\r\n", "\n").replace("\n", "\r\n")
    if body.startswith("Attribute VB_Name"):                    # .bas export: drop its own header line
        body = body.split("\r\n", 1)[1]
    return (head + body + ("\r\n" if not body.endswith("\r\n") else "")).encode("cp1252", "replace")


def _project_stream(mods):
    lines = [f'ID="{_ID}"']
    for name, kind, _ in mods:
        lines.append(f"Document={name}/&H00000000" if kind != "std" else f"Module={name}")
    lines += ['Name="VBAProject"', 'HelpContextID="0"', 'VersionCompatible32="393222000"',
              f'CMG="{_CMG}"', f'DPB="{_DPB}"', f'GC="{_GC}"', "",
              "[Host Extender Info]", "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000", "",
              "[Workspace]"]
    for name, kind, _ in mods:
        lines.append(f"{name}=0, 0, 0, 0, C" if kind != "std" else f"{name}=110, 145, 994, 721, ")
    return ("\r\n".join(lines) + "\r\n").encode("ascii")


def _projectwm(mods):
    b = bytearray()
    for name, _, _ in mods:
        b += name.encode("ascii") + b"\x00" + name.encode("utf-16-le") + b"\x00\x00"
    return bytes(b + b"\x00\x00")


# ── MS-CFB writer ──────────────────────────────────────────────────────────
class _Entry:
    def __init__(self, name, kind, data=b""):
        self.name, self.kind, self.data = name, kind, data       # kind: 5 root, 1 storage, 2 stream
        self.child = []
        self.left = self.right = self.sub = -1
        self.start, self.size = -2, len(data)
        self.sid = 0


def _key(e):
    return (len(e.name), e.name.upper())


def _link(children):
    """Balanced BST over name-sorted siblings → (root_sid, links set on entries)."""
    ch = sorted(children, key=_key)

    def build(lo, hi):
        if lo > hi:
            return -1
        mid = (lo + hi) // 2
        e = ch[mid]
        e.left = build(lo, mid - 1)
        e.right = build(mid + 1, hi)
        return e.sid
    return build(0, len(ch) - 1)


def build_cfb(streams):
    """streams: {'PROJECT': bytes, 'VBA/dir': bytes, ...}"""
    root = _Entry("Root Entry", 5)
    folders = {}
    for path, data in streams.items():
        parts = path.split("/")
        parent = root
        for p in parts[:-1]:
            if p not in folders:
                f = _Entry(p, 1)
                parent.child.append(f)
                folders[p] = f
            parent = folders[p]
        parent.child.append(_Entry(parts[-1], 2, data))
    entries = [root]

    def number(e):
        for c in e.child:
            c.sid = len(entries)
            entries.append(c)
        for c in e.child:
            number(c)
    number(root)
    for e in entries:
        if e.child:
            e.sub = _link(e.child)
    # mini stream for streams < 4096
    mini = bytearray()
    minifat = []
    big = []
    for e in entries[1:]:
        if e.kind != 2:
            continue
        if e.size < 4096:
            e.start = len(minifat) if e.size else -2
            nsec = (e.size + 63) // 64
            for i in range(nsec):
                minifat.append(len(minifat) + 1 if i < nsec - 1 else -2)
            mini += e.data + b"\x00" * (nsec * 64 - e.size)
        else:
            big.append(e)
    root.size = len(mini)
    # sector layout: [FAT..][DIR..][MINIFAT..][MINISTREAM..][big streams..]
    def nsec(nbytes):
        return (nbytes + 511) // 512
    n_dir = nsec(len(entries) * 128)
    n_minifat = nsec(len(minifat) * 4)
    n_mini = nsec(len(mini))
    n_big = sum(nsec(e.size) for e in big)
    n_fat = 1
    while True:
        total = n_fat + n_dir + n_minifat + n_mini + n_big
        need = (total + 127) // 128
        if need <= n_fat:
            break
        n_fat = need
    if n_fat > 109:
        raise ValueError("project too large for the single-DIFAT writer")
    fat = [-3] * n_fat                                          # FAT sectors marked 0xFFFFFFFD
    sectors = []                                                # payload sectors in order after FAT

    def chain(nb):
        first = n_fat + len(sectors)
        for i in range(nb):
            sectors.append(None)
            fat.append(first + i + 1 if i < nb - 1 else -2)
        return first if nb else -2
    dir_start = chain(n_dir)
    minifat_start = chain(n_minifat)
    root.start = chain(n_mini)
    payload = {}
    payload[dir_start] = None
    big_data = []
    for e in big:
        e.start = chain(nsec(e.size))
        big_data.append(e)
    fat += [-1] * (n_fat * 128 - len(fat))
    # serialise
    def dirent(e):
        nm = e.name.encode("utf-16-le") + b"\x00\x00"
        b = nm + b"\x00" * (64 - len(nm))
        b += struct.pack("<H", len(nm)) + bytes([e.kind, 1])
        b += struct.pack("<iii", e.left, e.right, e.sub)
        b += b"\x00" * 16 + struct.pack("<I", 0) + b"\x00" * 16
        b += struct.pack("<iI", e.start if e.size or e.kind == 5 else -2, e.size) + b"\x00\x00\x00\x00"
        return b
    dir_bytes = b"".join(dirent(e) for e in entries)
    dir_bytes += b"\x00" * (n_dir * 512 - len(dir_bytes))
    mf = b"".join(struct.pack("<i", v) for v in minifat)
    mf += b"\xff" * (n_minifat * 512 - len(mf))
    ms = bytes(mini) + b"\x00" * (n_mini * 512 - len(mini))
    body = dir_bytes + mf + ms
    for e in big_data:
        body += e.data + b"\x00" * (nsec(e.size) * 512 - e.size)
    header = bytearray(512)
    header[0:8] = bytes.fromhex("D0CF11E0A1B11AE1")
    struct.pack_into("<HHHHH", header, 24, 0x3E, 0x3, 0xFFFE, 9, 6)
    struct.pack_into("<IIiII", header, 40, 0, n_fat, dir_start, 0, 4096)
    struct.pack_into("<iIiI", header, 60, minifat_start if n_minifat else -2, n_minifat, -2, 0)
    for i in range(109):
        struct.pack_into("<i", header, 76 + 4 * i, i if i < n_fat else -1)
    fat_bytes = b"".join(struct.pack("<i", v) for v in fat)
    return bytes(header) + fat_bytes + body


def make_vba_project(modules):
    mods = [("ThisWorkbook", "workbook", "")] + list(modules)
    streams = {"PROJECT": _project_stream(mods), "PROJECTwm": _projectwm(mods),
               "VBA/_VBA_PROJECT": b"\xcc\x61\xff\xff\x01\x00\x00", "VBA/dir": _dir_stream(mods)}
    for name, kind, text in mods:
        streams["VBA/" + name] = compress(_module_source(name, kind, text))
    return build_cfb(streams)
