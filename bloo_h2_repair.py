"""
bloo_h2_repair.py
-----------------
Repairs Bloomsbury flat-file order imports where H2 (ship-to address) records
have been split across multiple physical lines due to embedded newlines in
address fields.

The format requires every record to be a single line of exactly 358 characters.
Bloomsbury's export system sometimes wraps long address fields, producing short
H2 lines followed by one or more continuation/blank lines.  The importer chokes
on the resulting blank lines with:

    Failure : begin 0, end 2, length 0

This script:
  1. Detects short H2 lines (< 358 chars)
  2. Absorbs all following blank / non-record continuation lines
  3. Joins them into a single 358-char record (pad or trim as needed)
  4. Drops orphaned blank lines elsewhere in the file
  5. Rewrites the EOF line count to match the repaired line total
  6. Writes a new output file alongside the original (suffix _repaired)

Usage:
    python bloo_h2_repair.py <input_file> [output_file]

    If output_file is omitted, the repaired file is written next to the input
    with _repaired inserted before the extension.

"""

import sys
import os
import re

# ── Constants ─────────────────────────────────────────────────────────────────

VALID_PREFIXES = ('H1', 'H2', 'H3', 'D1', 'D2', '$$')
H2_TARGET_LEN  = 358   # Fixed-width length every H2 line must be


# ── Helpers ───────────────────────────────────────────────────────────────────

def is_record_line(line: str) -> bool:
    """Return True if the line starts with a known record-type prefix."""
    return any(line.startswith(p) for p in VALID_PREFIXES)


def update_eof_count(eof_line: str, new_count: int) -> str:
    """Replace the 7-digit trailing count in the $$EOF line."""
    # Format: $$EOFxxxx  NNNNNNN   YYYYMMDDHHMMSSxxxxxxx
    # The count occupies the last 7 characters.
    return eof_line[:-7] + f"{new_count:07d}"


# ── Core repair ───────────────────────────────────────────────────────────────

def repair(lines: list[str]) -> tuple[list[str], list[dict]]:
    """
    Process lines, collapsing broken H2 records.

    Returns:
        (output_lines, repair_log)
        repair_log is a list of dicts describing each repaired record.
    """
    output   = []
    log      = []
    i        = 0

    while i < len(lines):
        line = lines[i].rstrip('\r')

        # ── Blank line outside an H2 collapse → discard ──────────────────────
        if line == '':
            i += 1
            continue

        # ── Broken H2: shorter than expected ─────────────────────────────────
        if line.startswith('H2') and len(line) < H2_TARGET_LEN:
            order_ref = line[2:17].strip()
            parts     = [line]
            j         = i + 1

            # Absorb continuation lines until the next proper record
            while j < len(lines):
                cont = lines[j].rstrip('\r')
                if is_record_line(cont):
                    break          # Next real record — stop here
                parts.append(cont) # Blank or continuation — absorb
                j += 1

            joined   = ''.join(parts)
            repaired = joined.ljust(H2_TARGET_LEN)[:H2_TARGET_LEN]
            output.append(repaired)

            log.append({
                'order':          order_ref,
                'input_line':     i + 1,
                'lines_consumed': j - i,
                'joined_len':     len(joined),
                'action':         'padded' if len(joined) <= H2_TARGET_LEN else 'trimmed',
            })
            i = j

        # ── Normal line ───────────────────────────────────────────────────────
        else:
            output.append(line)
            i += 1

    return output, log


# ── EOF count fixup ───────────────────────────────────────────────────────────

def fix_eof(output: list[str]) -> list[str]:
    """Update the $$EOF line's trailing count to match repaired line count."""
    # Count = total lines minus the $$HDR and $$EOF lines themselves
    data_lines = sum(1 for l in output if not l.startswith('$$'))
    result     = []
    for line in output:
        if line.startswith('$$EOF'):
            line = update_eof_count(line, data_lines)
        result.append(line)
    return result


# ── I/O ───────────────────────────────────────────────────────────────────────

def derive_output_path(input_path: str) -> str:
    root, ext = os.path.splitext(input_path)
    return root + '_repaired' + ext


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    input_path  = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else derive_output_path(input_path)

    if not os.path.isfile(input_path):
        print(f"ERROR: File not found: {input_path}")
        sys.exit(1)

    # Read
    with open(input_path, 'r', encoding='utf-8') as fh:
        raw_lines = fh.read().split('\n')

    print(f"Input  : {input_path}")
    print(f"Lines  : {len(raw_lines)} (raw)")

    # Repair H2 records
    output, log = repair(raw_lines)

    # Fix EOF count
    output = fix_eof(output)

    # Write
    with open(output_path, 'w', encoding='utf-8', newline='\n') as fh:
        fh.write('\n'.join(output) + '\n')

    # Report
    print(f"Output : {output_path}")
    print(f"Lines  : {len(output)} (repaired)")
    print()

    if log:
        print(f"H2 records repaired: {len(log)}")
        print()
        col = "{:<16} {:>12} {:>16} {:>12} {:>8}"
        print(col.format("Order ref", "Input line", "Lines consumed", "Joined len", "Action"))
        print("-" * 68)
        for r in log:
            print(col.format(
                r['order'],
                r['input_line'],
                r['lines_consumed'],
                r['joined_len'],
                r['action'],
            ))
    else:
        print("No broken H2 records found — file was already clean.")

    print()

    # Sanity checks
    blanks     = [l for l in output if l == '']
    short_h2s  = [l for l in output if l.startswith('H2') and len(l) < H2_TARGET_LEN]
    eof_lines  = [l for l in output if l.startswith('$$EOF')]

    print("── Sanity checks ──────────────────────────────────────")
    print(f"  Blank lines remaining : {len(blanks)}"  + ("  ✓" if not blanks  else "  ✗ PROBLEM"))
    print(f"  Short H2s remaining   : {len(short_h2s)}" + ("  ✓" if not short_h2s else "  ✗ PROBLEM"))
    if eof_lines:
        declared = int(eof_lines[0][-7:])
        actual   = sum(1 for l in output if not l.startswith('$$'))
        match    = declared == actual
        print(f"  EOF count ({declared}) vs data lines ({actual}) : " + ("✓" if match else "✗ MISMATCH"))
    print()


if __name__ == '__main__':
    main()
