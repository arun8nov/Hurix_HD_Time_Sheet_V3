import os
import re

def check_js_balance(content, label):
    stack = []
    pairs = {')': '(', ']': '[', '}': '{'}
    line = 1
    col = 1
    in_single_quote = False
    in_double_quote = False
    in_backtick = False
    in_single_comment = False
    in_multi_comment = False
    escape = False

    i = 0
    n = len(content)
    while i < n:
        ch = content[i]
        if ch == '\n':
            line += 1
            col = 1
            if in_single_comment:
                in_single_comment = False
            i += 1
            continue
        col += 1
        
        if escape:
            escape = False
            i += 1
            continue

        if ch == '\\':
            if in_single_quote or in_double_quote or in_backtick:
                escape = True
            i += 1
            continue

        if in_single_comment:
            i += 1
            continue

        if in_multi_comment:
            if ch == '*' and i + 1 < n and content[i+1] == '/':
                in_multi_comment = False
                i += 2
                continue
            i += 1
            continue

        if in_single_quote:
            if ch == "'":
                in_single_quote = False
            i += 1
            continue

        if in_double_quote:
            if ch == '"':
                in_double_quote = False
            i += 1
            continue

        if in_backtick:
            if ch == '`':
                in_backtick = False
            i += 1
            continue

        if ch == '/' and i + 1 < n:
            if content[i+1] == '/':
                in_single_comment = True
                i += 2
                continue
            elif content[i+1] == '*':
                in_multi_comment = True
                i += 2
                continue

        if ch == "'":
            in_single_quote = True
        elif ch == '"':
            in_double_quote = True
        elif ch == '`':
            in_backtick = True
        elif ch in '([{':
            stack.append((ch, line, col))
        elif ch in ')]}':
            if not stack:
                print(f'Error in {label}: Unmatched closing {ch} at line {line}, col {col}')
                return False
            top, tline, tcol = stack.pop()
            if pairs[ch] != top:
                print(f'Error in {label}: Mismatched {top} from line {tline} with {ch} at line {line}')
                return False
        i += 1

    if stack:
        print(f'Error in {label}: Unclosed {len(stack)} delimiters! First: {stack[0]}')
        return False
    print(f'{label}: Delimiters balanced successfully!')
    return True

# 1. Check Code.gs (entire file is JS)
check_js_balance(open('Code.gs', 'r', encoding='utf-8').read(), 'Code.gs')

# 2. Check script tags in Index.html
idx_content = open('Index.html', 'r', encoding='utf-8').read()
script_matches = list(re.finditer(r'<script\b[^>]*>(.*?)</script>', idx_content, re.DOTALL | re.IGNORECASE))
print(f'Found {len(script_matches)} <script> blocks in Index.html')
for s_idx, match in enumerate(script_matches):
    check_js_balance(match.group(1), f'Index.html <script #{s_idx+1}>')

# 3. Check HTML tag balance in AdminView.html
admin_content = open('AdminView.html', 'r', encoding='utf-8').read()
# Check critical IDs
inspection_ids = [
    'inspectionEmpName',
    'inspectionEmpIdTeam',
    'inspectionSummaryBadges',
    'inspectionWorkLogsTableBody',
    'inspectionWorkLogsTableFoot',
    'empWorkLogsInspectionModal',
    'adminMasterSheetLink',
    'adminSearchInput',
    'adminRefSummaryTableBody',
    'adminRefSummaryTableFoot',
    'adminStaffTableBody',
    'adminFilterClient',
    'adminFilterProject',
    'adminFilterShift',
    'adminFilterTeam'
]

print('\nChecking critical Admin IDs in AdminView.html:')
all_pass = True
for elem_id in inspection_ids:
    if f'id="{elem_id}"' in admin_content:
        print(f'  [PASS] id="{elem_id}"')
    else:
        print(f'  [FAIL] id="{elem_id}"')
        all_pass = False

if all_pass:
    print('\nALL ADMIN AUDIT CHECKS PASSED!')
