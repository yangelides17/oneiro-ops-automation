#!/usr/bin/env python3
"""
Systematically migrate ALL old field name references in React components
to match the new Drizzle camelCase API response shapes.

Safe to run multiple times — idempotent.
"""
import os
import re
import glob

SRC_DIR = os.path.join(os.path.dirname(__file__), 'src')

# ── All dot-access field renames ──────────────────────────────
# Pattern: .old_name → .newName
# These are applied as simple string replacements on dot-access patterns.
DOT_RENAMES = {
    '.contract_num':          '.contractNum',
    '.due_date':              '.dueDate',
    '.work_type':             '.workType',
    '.work_end':              '.workEndDate',
    '.work_start':            '.workStartDate',
    '.from_street':           '.fromStreet',
    '.to_street':             '.toStreet',
    '.quantity_unit':         '.quantityUnit',
    '.markings_total':        '.markingsTotal',
    '.markings_completed':    '.markingsCompleted',
    '.geocode_warning':       '.geocodeWarning',
    '.color_material':        '.colorMaterial',
    '.date_completed':        '.dateCompleted',
    '.wo_section':            '.woSection',
    '.crew_chief':            '.crewChief',
    '.doc_type':              '.docType',
    '.added_by':              '.addedBy',
    '.time_in':               '.timeIn',
    '.time_out':              '.timeOut',
    '.dispatch_date':         '.dispatchDate',
    '.wo_received':           '.woReceivedDate',
    '.admin_reviewed':        '.adminReviewed',
    '.review_notes':          '.reviewNotes',
    '.contract_id':           '.contractId',
    '.wo_id':                 '.woId',
    '.scan_file_id':          '.scanFileKey',
    '.original_filename':     '.originalFilename',
    '.date_entered':          '.dateEntered',
    '.general_remarks':       '.generalRemarks',
    '.water_blast_confirmed': '.waterBlastConfirmed',
    '.folder_url':            '.folderUrl',
    '.file_id':               '.fileId',
    '.by_contractor':         '.byContractor',
    '.by_group':              '.byGroup',
    '.labor_daily':           '.laborDaily',
    '.labor_totals':          '.laborTotals',
    '.by_doc_type':           '.byDocType',
    '.prime_contractor':      '.contractorName',
}

# Object key renames (in JSON bodies, destructuring, object literals)
KEY_RENAMES = {
    'color_material:':    'colorMaterial:',
    'work_type:':         'workType:',
    'date_completed:':    'dateCompleted:',
    'added_by:':          'addedBy:',
    'crew_chief:':        'crewChief:',
    'wo_id:':             'woId:',
    'contract_number:':   'contractNum:',
    'doc_types:':         'docTypes:',
    'time_in:':           'timeIn:',
    'time_out:':          'timeOut:',
    "'color_material'":   "'colorMaterial'",
    "'work_type'":        "'workType'",
    "'wo_id'":            "'woId'",
}

# Stats field renames
STATS_RENAMES = {
    'stats.complete,':    'stats.completed,',
    'stats.complete ':    'stats.completed ',
    'stats.complete}':    'stats.completed}',
    'stats?.complete,':   'stats?.completed,',
    'stats?.complete ':   'stats?.completed ',
    'stats?.complete}':   'stats?.completed}',
    'stats.in_progress':  'stats.inProgress',
    'stats?.in_progress': 'stats?.inProgress',
}

# Sign-in specific field renames
SIGNIN_RENAMES = {
    'bill_contract_number': 'contractNum',
    'bill_borough':         'regionCode',
    'contractNumber':       'contractNum',  # old camelCase variant
    'prime_contractor':     'contractorName',
}

# Revenue/dashboard response field renames
RESPONSE_RENAMES = {
    'unpriced_items':     'needsPricing',
    'invoiced_revenue':   'invoicedRevenue',
    'wip_revenue':        'wipRevenue',
}

# API path fixes
API_PATH_RENAMES = {
    '/api/wo-markings/': '/api/wos/',  # Note: needs manual suffix fix
}


def apply_regex_renames(content):
    """Apply regex-based renames that need word boundary matching."""

    # .borough → .regionCode (only field access, not "Borough" label text)
    content = re.sub(r'(\w)\.borough(?=[^a-zA-Z]|$)', r'\1.regionCode', content)

    # .contractor → .contractorName (only field access, not "Contractor" label)
    content = re.sub(r'(\w)\.contractor(?!Name)(?=[^a-zA-Z]|$)', r'\1.contractorName', content)

    # .water_blast → .waterBlastRequired (but NOT .water_blast_confirmed which is already handled)
    content = re.sub(r'\.water_blast(?!_confirmed|Required)', '.waterBlastRequired', content)

    # stats?.complete → stats?.completed
    content = re.sub(r'stats\?\.complete(?!d)', 'stats?.completed', content)
    content = re.sub(r'stats\.complete(?!d)', 'stats.completed', content)

    return content


def apply_wo_id_renames(content, filename):
    """
    Replace wo.id with wo.woNumber where it's used as a display value.

    This is the trickiest rename: wo.id in the old system WAS the WO number
    (e.g. "RM-43402"). In the new system, wo.id is a UUID and wo.woNumber
    is the WO number. We need to change display/search/filter uses but NOT
    React key= props (which can use any unique value including UUID).
    """

    # Only apply to files that deal with work order display
    wo_files = [
        'Dashboard.jsx', 'FieldReport.jsx', 'NavTab.jsx', 'ScanWO.jsx',
        'DeleteWOModal.jsx', 'WODocsQueue.jsx', 'InvoiceCell.jsx',
        'StatusPickerModal.jsx', 'DocStatusChips.jsx',
    ]
    if not any(filename.endswith(f) for f in wo_files):
        return content

    # wo.id in template strings used for display/links (NOT key={wo.id})
    # Pattern: ${wo.id} or ${encodeURIComponent(wo.id)} in template literals
    content = content.replace('${wo.id}', '${wo.woNumber}')
    content = content.replace('${encodeURIComponent(wo.id)}', '${encodeURIComponent(wo.woNumber)}')

    # wo.id in JSX text content: {wo.id} for display
    content = content.replace('{wo.id}', '{wo.woNumber}')

    # wo.id in search/filter: wo.id.toLowerCase()
    content = content.replace('wo.id.toLowerCase', 'wo.woNumber.toLowerCase')
    content = content.replace('wo.id.toUpperCase', 'wo.woNumber.toUpperCase')

    # wo.id in string concatenation for labels
    content = re.sub(r'`\$\{wo\.id\}', '`${wo.woNumber}', content)

    # attention.includes(wo.id) — filtering by WO number
    content = content.replace('attention.includes(wo.id)', 'attention.includes(wo.woNumber)')

    # WO dropdown label: wo.id — location
    content = re.sub(r'wo\.id\s*—', 'wo.woNumber —', content)
    # Search string: wo.id (space) in concatenated search
    content = re.sub(r'`\$\{wo\.id\}\s', '`${wo.woNumber} ', content)

    return content

    return content


def migrate_file(filepath):
    """Apply all renames to a single file."""
    with open(filepath, 'r', encoding='utf-8') as f:
        original = f.read()

    content = original

    # 1. Dot-access renames
    for old, new in DOT_RENAMES.items():
        content = content.replace(old, new)

    # 2. Object key renames
    for old, new in KEY_RENAMES.items():
        content = content.replace(old, new)

    # 3. Stats renames
    for old, new in STATS_RENAMES.items():
        content = content.replace(old, new)

    # 4. Sign-in specific
    for old, new in SIGNIN_RENAMES.items():
        content = content.replace(old, new)

    # 5. Revenue/dashboard response
    for old, new in RESPONSE_RENAMES.items():
        content = content.replace(old, new)

    # 6. Regex-based renames
    content = apply_regex_renames(content)

    # 7. wo.id → wo.woNumber for display contexts
    filename = os.path.basename(filepath)
    content = apply_wo_id_renames(content, filename)

    # Only write if changed
    if content != original:
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(content)
        changes = sum(1 for a, b in zip(original.split('\n'), content.split('\n')) if a != b)
        return changes
    return 0


def main():
    # Find all .jsx and .js files
    patterns = [
        os.path.join(SRC_DIR, '**', '*.jsx'),
        os.path.join(SRC_DIR, '**', '*.js'),
    ]

    files = []
    for pattern in patterns:
        files.extend(glob.glob(pattern, recursive=True))

    total_changes = 0
    for filepath in sorted(files):
        rel = os.path.relpath(filepath, SRC_DIR)
        changes = migrate_file(filepath)
        if changes > 0:
            print(f'  {rel}: {changes} lines changed')
            total_changes += changes

    print(f'\nTotal: {total_changes} lines changed across {len(files)} files')

    # Verify no file was emptied
    for filepath in files:
        if os.path.getsize(filepath) == 0:
            print(f'  ERROR: {filepath} is empty!')


if __name__ == '__main__':
    main()
