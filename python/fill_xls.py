#!/usr/bin/env python3
"""
fill_xls.py - Fills XLS template by patching the BIFF8 binary stream.

Approach:
  1. Scan globals section for BOUNDSHEET records to get lbPlyPos of every sheet
  2. Identify which BOUNDSHEETs are worksheets by reading the BOF dt field at
     lbPlyPos (avoids BIFF-version ambiguity in the BOUNDSHEET dt byte offset)
  3. Depth-track from target sheet's BOF to find its matching EOF
  4. Rebuild the sheet block: remove INDEX/DBCELL, update DIMENSIONS, inject rows
  5. Fix lbPlyPos in ALL BOUNDSHEET records for sections after the patched sheet

Usage:
  python3 fill_xls.py <config.json>
"""

import sys
import json
import struct
import xlrd
import olefile
from xlwt.CompoundDoc import XlsDoc

# BIFF8 record type codes
BIFF_BOF        = 0x0809
BIFF_EOF        = 0x000A
BIFF_BOUNDSHEET = 0x0085
BIFF_INDEX      = 0x020B
BIFF_DBCELL     = 0x00D7
BIFF_DIMENSIONS = 0x0200
BIFF_ROW        = 0x0208
BIFF_LABEL      = 0x0204
BIFF_NUMBER     = 0x0203

# BOF document-type codes (2-byte dt field inside the BOF payload)
BOF_DT_GLOBALS   = 0x0005
BOF_DT_WORKSHEET = 0x0010


# ── OLE2 helpers ──────────────────────────────────────────────────────────────

def read_biff_stream(path):
    ole = olefile.OleFileIO(path)
    try:
        for name in ('Workbook', 'Book'):
            if ole.exists(name):
                return ole.openstream(name).read()
        raise ValueError("No Workbook/Book stream in: " + path)
    finally:
        ole.close()


# ── BIFF record building ──────────────────────────────────────────────────────

def pack_biff(type_code, data):
    MAX = 8224
    if len(data) <= MAX:
        return struct.pack('<HH', type_code, len(data)) + data
    result = struct.pack('<HH', type_code, MAX) + data[:MAX]
    rest = data[MAX:]
    while rest:
        chunk = rest[:MAX]
        result += struct.pack('<HH', 0x003C, len(chunk)) + chunk
        rest = rest[len(chunk):]
    return result


def make_row_record(row_num, first_col, last_col):
    return struct.pack('<HHHHHHHH',
        row_num, first_col, last_col + 1, 0xFF, 0, 0, 0x0100, 0x0F)


def make_label_record(row, col, xf_idx, text):
    s = str(text)
    try:
        encoded = s.encode('latin-1')
        flags = 0x00
    except (UnicodeEncodeError, UnicodeDecodeError):
        encoded = s.encode('utf-16-le')
        flags = 0x01
    return struct.pack('<HHHH', row, col, xf_idx, len(s)) + bytes([flags]) + encoded


def make_number_record(row, col, xf_idx, value):
    return struct.pack('<HHHd', row, col, xf_idx, float(value))


# ── BIFF structure analysis ───────────────────────────────────────────────────

def bof_dt(biff_data, bof_pos):
    """
    Return the document-type (dt) field from a BOF record.
    BOF payload layout: vers(2) + dt(2) + ...
    Returns 0 if bof_pos doesn't point at a valid BOF.
    """
    if bof_pos + 8 > len(biff_data):
        return 0
    tc = struct.unpack_from('<H', biff_data, bof_pos)[0]
    if tc != BIFF_BOF:
        return 0
    return struct.unpack_from('<H', biff_data, bof_pos + 6)[0]  # pos+4 data + 2 vers


def parse_boundsheets(biff_data):
    """
    Scan the globals section and collect all BOUNDSHEET records.
    Returns list of {'pos': int, 'lbPlyPos': int} — the dt field is NOT read
    here; use bof_dt(biff_data, lbPlyPos) instead to avoid version ambiguity.

    Key fix: use a `globals_done` boolean so that after the globals EOF,
    subsequent worksheet BOFs (depth 0→1) do NOT re-trigger globals mode.
    """
    results      = []
    pos          = 0
    depth        = 0
    in_globals   = False
    globals_done = False

    while pos + 4 <= len(biff_data):
        tc  = struct.unpack_from('<H', biff_data, pos)[0]
        rln = struct.unpack_from('<H', biff_data, pos + 2)[0]

        if tc == BIFF_BOF:
            depth += 1
            # Only enter globals mode on the very first BOF in the stream
            if depth == 1 and not globals_done:
                in_globals = True

        elif tc == BIFF_EOF:
            if depth == 1 and in_globals:
                in_globals   = False
                globals_done = True
            if depth > 0:
                depth -= 1

        elif tc == BIFF_BOUNDSHEET and in_globals and rln >= 6:
            lbp = struct.unpack_from('<I', biff_data, pos + 4)[0]
            results.append({'pos': pos, 'lbPlyPos': lbp})

        pos += 4 + rln

        if globals_done:
            break   # All BOUNDSHEET records have been collected

    return results


def find_sheet_eof(biff_data, bof_pos):
    """
    Starting at bof_pos (a BOF record), scan forward with a depth counter
    and return the byte offset of the matching EOF record.
    Handles nested chart/VBA BOF+EOF pairs inside the sheet correctly.
    """
    pos   = bof_pos
    depth = 0

    while pos + 4 <= len(biff_data):
        tc  = struct.unpack_from('<H', biff_data, pos)[0]
        rln = struct.unpack_from('<H', biff_data, pos + 2)[0]

        if tc == BIFF_BOF:
            depth += 1
        elif tc == BIFF_EOF:
            depth -= 1
            if depth == 0:
                return pos

        pos += 4 + rln

    raise RuntimeError(f"No matching EOF found for BOF at offset {bof_pos}")


# ── Core patch ────────────────────────────────────────────────────────────────

def rebuild_sheet_block(sheet_bytes, new_row_biff, total_new_rows, header_row):
    """
    Rebuild one raw worksheet block (BOF…EOF inclusive):
      - Remove INDEX records (stale absolute offsets)
      - Remove DBCELL records (referenced only by INDEX)
      - Update DIMENSIONS to cover new rows
      - Inject new_row_biff just before the outermost EOF
    Uses depth tracking so nested chart records inside the sheet are untouched.
    """
    rebuilt = bytearray()
    pos     = 0
    depth   = 0

    while pos + 4 <= len(sheet_bytes):
        tc  = struct.unpack_from('<H', sheet_bytes, pos)[0]
        rln = struct.unpack_from('<H', sheet_bytes, pos + 2)[0]
        rec = sheet_bytes[pos + 4 : pos + 4 + rln]

        if tc == BIFF_BOF:
            depth += 1
            rebuilt += sheet_bytes[pos : pos + 4 + rln]

        elif tc == BIFF_EOF:
            depth -= 1
            if depth == 0:
                rebuilt += bytes(new_row_biff)   # inject before outermost EOF
            rebuilt += sheet_bytes[pos : pos + 4 + rln]

        elif depth == 1 and tc == BIFF_INDEX:
            print(f"DEBUG removed INDEX ({rln} bytes)", file=sys.stderr)

        elif depth == 1 and tc == BIFF_DBCELL:
            pass   # remove silently

        elif depth == 1 and tc == BIFF_DIMENSIONS and rln == 14:
            # BIFF8 DIMENSIONS (14 bytes):
            # first_row(4) last_row+1(4) first_col(2) last_col+1(2) reserved(2)
            first_row = struct.unpack_from('<I', rec, 0)[0]
            old_lr1   = struct.unpack_from('<I', rec, 4)[0]
            first_col = struct.unpack_from('<H', rec, 8)[0]
            last_col1 = struct.unpack_from('<H', rec, 10)[0]
            reserved  = struct.unpack_from('<H', rec, 12)[0]
            new_lr1   = max(old_lr1, header_row + total_new_rows + 1)
            new_rec   = struct.pack('<IIHHH',
                first_row, new_lr1, first_col, last_col1, reserved)
            rebuilt  += struct.pack('<HH', tc, len(new_rec)) + new_rec
            print(f"DEBUG DIMENSIONS last_row updated to {new_lr1 - 1}", file=sys.stderr)

        else:
            rebuilt += sheet_bytes[pos : pos + 4 + rln]

        pos += 4 + rln

    return bytes(rebuilt)


def patch_biff(biff_data, xlrd_sheet_idx, new_row_biff, total_new_rows, header_row):
    """
    Full patch pipeline:
      1. Parse BOUNDSHEET records from globals section
      2. Filter to worksheet BOFs by reading the BOF dt field at each lbPlyPos
         (avoids BOUNDSHEET.dt byte-offset ambiguity across BIFF versions)
      3. Depth-scan from target BOF to find its matching EOF
      4. Rebuild the target sheet block
      5. Fix lbPlyPos in ALL subsequent BOUNDSHEET records
    """
    all_bs = parse_boundsheets(biff_data)
    print(f"DEBUG total BOUNDSHEET records={len(all_bs)}", file=sys.stderr)

    # Filter to worksheet-type BOFs by reading the BOF at lbPlyPos
    ws_bs = [b for b in all_bs if bof_dt(biff_data, b['lbPlyPos']) == BOF_DT_WORKSHEET]
    print(f"DEBUG worksheet BOUNDSHEETs={len(ws_bs)} "
          f"all_dts={[hex(bof_dt(biff_data, b['lbPlyPos'])) for b in all_bs[:5]]}",
          file=sys.stderr)

    if xlrd_sheet_idx >= len(ws_bs):
        raise RuntimeError(
            f"xlrd_sheet_idx={xlrd_sheet_idx} but only {len(ws_bs)} "
            f"worksheet BOUNDSHEET records found (total={len(all_bs)})"
        )

    target_bs  = ws_bs[xlrd_sheet_idx]
    bof_pos    = target_bs['lbPlyPos']
    eof_pos    = find_sheet_eof(biff_data, bof_pos)
    sheet_end  = eof_pos + 4   # 4 bytes: type(2) + len(2), EOF has no payload

    print(f"DEBUG target sheet[{xlrd_sheet_idx}] BOF={bof_pos} EOF={eof_pos} "
          f"ws_bs_count={len(ws_bs)} all_bs_count={len(all_bs)}",
          file=sys.stderr)

    old_block = biff_data[bof_pos : sheet_end]
    new_block = rebuild_sheet_block(old_block, new_row_biff, total_new_rows, header_row)
    delta     = len(new_block) - len(old_block)

    print(f"DEBUG sheet block: {len(old_block)} → {len(new_block)} (delta={delta:+d})",
          file=sys.stderr)

    # Assemble modified BIFF
    modified = bytearray(biff_data[:bof_pos] + new_block + biff_data[sheet_end:])

    # Fix ALL BOUNDSHEET offsets for sections that start at or after sheet_end
    for bs in all_bs:
        if bs['lbPlyPos'] >= sheet_end:
            new_off = bs['lbPlyPos'] + delta
            struct.pack_into('<I', modified, bs['pos'] + 4, new_off)
            print(f"DEBUG BOUNDSHEET lbPlyPos {bs['lbPlyPos']} → {new_off}",
                  file=sys.stderr)

    return bytes(modified)


# ── xlrd data helpers ─────────────────────────────────────────────────────────

def get_col_map(sheet, header_row):
    col_map = {}
    for c in range(sheet.ncols):
        raw = sheet.cell_value(header_row, c)
        key = str(raw).strip() if raw else ''
        if key:
            col_map[key] = c
    return col_map


def get_data_rows(sheet, header_row, col_map):
    rows = []
    for r in range(header_row + 1, sheet.nrows):
        row      = {}
        has_data = False
        for name, c in col_map.items():
            ct = sheet.cell_type(r, c)
            v  = sheet.cell_value(r, c)
            if ct == xlrd.XL_CELL_EMPTY:
                row[name] = None
            elif ct == xlrd.XL_CELL_TEXT:
                s = str(v).strip()
                row[name] = s if s else None
                if s:
                    has_data = True
            elif ct == xlrd.XL_CELL_NUMBER:
                row[name] = int(v) if v == int(v) else v
                has_data   = True
            elif ct in (xlrd.XL_CELL_DATE, xlrd.XL_CELL_BOOLEAN):
                row[name] = v
                has_data   = True
            else:
                s = str(v).strip() if v else None
                row[name] = s
                if s:
                    has_data = True
        if has_data:
            rows.append(row)
    return rows


# ── main ──────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print("Usage: fill_xls.py <config.json>", file=sys.stderr)
        sys.exit(1)

    with open(sys.argv[1], 'r', encoding='utf-8') as f:
        cfg = json.load(f)

    template_path   = cfg['templatePath']
    old_path        = cfg['oldPath']
    output_path     = cfg['outputPath']
    mappings_path   = cfg['mappingsPath']
    tmpl_sheet_idx  = int(cfg.get('templateSheetIdx',  0))
    tmpl_header_row = int(cfg.get('templateHeaderRow', 0))
    old_sheet_idx   = int(cfg.get('oldSheetIdx',       0))
    old_header_row  = int(cfg.get('oldHeaderRow',      0))

    with open(mappings_path, 'r', encoding='utf-8') as f:
        mappings = json.load(f)

    # ── Template column map ───────────────────────────────────────────────────
    tmpl_wb = xlrd.open_workbook(template_path)
    if tmpl_sheet_idx >= tmpl_wb.nsheets:
        tmpl_sheet_idx = 0
    tmpl_sheet   = tmpl_wb.sheet_by_index(tmpl_sheet_idx)
    tmpl_col_map = get_col_map(tmpl_sheet, tmpl_header_row)
    print(f"DEBUG template '{tmpl_wb.sheet_names()[tmpl_sheet_idx]}' "
          f"hdr={tmpl_header_row} cols={list(tmpl_col_map.keys())[:6]}",
          file=sys.stderr)

    # ── Data rows — use pre-processed rows from frontend if provided ──────────
    # The frontend may send dataRows that already have prefix/suffix applied.
    # Use those directly to preserve all transformations (prefix, suffix, edits).
    pre_data_rows = cfg.get('dataRows')
    if pre_data_rows is not None:
        # Convert column keys: frontend uses newCol names, Python also uses newCol,
        # so they match directly through mappings.
        data_rows = pre_data_rows
        print(f"DEBUG using pre-processed dataRows from frontend: {len(data_rows)} rows",
              file=sys.stderr)
    else:
        # Fallback: read from old catalog file directly
        old_wb = xlrd.open_workbook(old_path)
        if old_sheet_idx >= old_wb.nsheets:
            old_sheet_idx = 0
        old_sheet   = old_wb.sheet_by_index(old_sheet_idx)
        old_col_map = get_col_map(old_sheet, old_header_row)
        data_rows   = get_data_rows(old_sheet, old_header_row, old_col_map)
        print(f"DEBUG old '{old_wb.sheet_names()[old_sheet_idx]}' "
              f"hdr={old_header_row} data_rows={len(data_rows)}",
              file=sys.stderr)

    # ── Build new BIFF row records ────────────────────────────────────────────
    new_row_biff  = bytearray()
    cells_written = 0
    rows_written  = 0

    for row_idx, row_data in enumerate(data_rows):
        target_row = tmpl_header_row + 1 + row_idx
        cells      = []

        for m in mappings:
            new_col = (m.get('newCol') or '').strip()
            old_col = (m.get('oldCol') or '').strip()
            if not new_col or not old_col:
                continue
            col_idx = tmpl_col_map.get(new_col)
            if col_idx is None:
                continue
            value = row_data.get(old_col)
            if value is None:
                continue
            cells.append((col_idx, value))

        if not cells:
            continue

        first_col = min(c for c, _ in cells)
        last_col  = max(c for c, _ in cells)
        new_row_biff += pack_biff(
            BIFF_ROW, make_row_record(target_row, first_col, last_col))

        for col_idx, value in sorted(cells, key=lambda x: x[0]):
            if isinstance(value, (int, float)):
                new_row_biff += pack_biff(
                    BIFF_NUMBER, make_number_record(target_row, col_idx, 0, value))
            else:
                new_row_biff += pack_biff(
                    BIFF_LABEL, make_label_record(target_row, col_idx, 0, value))
            cells_written += 1

        rows_written += 1

    print(f"DEBUG cells_written={cells_written} rows_written={rows_written}",
          file=sys.stderr)

    # ── Read BIFF, patch, write ───────────────────────────────────────────────
    biff_data = read_biff_stream(template_path)
    print(f"DEBUG original BIFF={len(biff_data)} bytes", file=sys.stderr)

    modified = patch_biff(
        biff_data, tmpl_sheet_idx, new_row_biff, rows_written, tmpl_header_row)
    print(f"DEBUG modified BIFF={len(modified)} bytes", file=sys.stderr)

    doc = XlsDoc()
    doc.save(output_path, modified)

    # ── Verify with olefile + xlrd ────────────────────────────────────────────
    try:
        vbiff = read_biff_stream(output_path)
        print(f"DEBUG OLE2 OK: {len(vbiff)} bytes", file=sys.stderr)
        vwb = xlrd.open_workbook(output_path)
        vs  = vwb.sheet_by_index(tmpl_sheet_idx)
        print(f"DEBUG xlrd OK: {vwb.nsheets} sheets, "
              f"target nrows={vs.nrows} ncols={vs.ncols}",
              file=sys.stderr)
    except Exception as e:
        print(f"DEBUG verify FAILED: {e}", file=sys.stderr)
        raise

    print(json.dumps({
        "success":      True,
        "rowsWritten":  rows_written,
        "cellsWritten": cells_written,
    }))


if __name__ == '__main__':
    main()
