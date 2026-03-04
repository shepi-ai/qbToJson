import re

def parse_journal_line(line):
    """Parse a single journal entry line from pdfplumber format"""
    #  Example: '05/14/2025 Invoice 1027 Bill's Windsurf Shop Accounts Receivable (A/R) 85.00'
    
    # Extract amounts from end (last 1 or 2 numbers)
    amounts = re.findall(r'[\d,]+\.\d{2}$|[\d,]+$', line)
    if not amounts:
        return None
    
    # Remove amounts from line to get rest
    line_without_amounts = line
    for amt in amounts:
        line_without_amounts = line_without_amounts.rsplit(amt, 1)[0].strip()
    
    # Extract date if present
    date_match = re.match(r'^(\d{2}/\d{2}/\d{4})\s+(.+)$', line_without_amounts)
    
    if date_match:
        date = date_match.group(1)
        rest = date_match.group(2).strip()
        
        # Parse rest: TYPE NUM NAME... ACCOUNT
        # Account is the last part before amounts
        parts = rest.split()
        
        trans_type = parts[0] if len(parts) > 0 else ''
        num = parts[1] if len(parts) > 1 and parts[1].isdigit() else ''
        
        # Everything else is name/memo/account - account is last part
        remaining_idx = 2 if num else 1
        remaining = ' '.join(parts[remaining_idx:])
        
        # Account is typically the last "word" or phrase
        # For simplicity, take last 3-5 words as account
        account_words = remaining.split()[-5:]  # Adjust as needed
        account = ' '.join(account_words)
        
        # Parse amounts
        if len(amounts) == 2:
            debit = float(amounts[0].replace(',', ''))
            credit = float(amounts[1].replace(',', ''))
        else:
            debit = 0.0
            credit = float(amounts[0].replace(',', ''))
        
        return {
            'date': date,
            'type': trans_type,
            'num': num,
            'account': account,
            'debit': debit,
            'credit': credit
        }
    else:
        # No date - continuation line
        # Format: MEMO... ACCOUNT AMOUNT
        # Account is last part before amounts
        words = line_without_amounts.split()
        account = ' '.join(words[-3:]) if len(words) >= 3 else line_without_amounts
        
        if len(amounts) == 2:
            debit = float(amounts[0].replace(',', ''))
            credit = float(amounts[1].replace(',', ''))
        else:
            debit = 0.0
            credit = float(amounts[0].replace(',', ''))
        
        return {
            'account': account,
            'debit': debit,
            'credit': credit
        }

# Test
line = '05/14/2025 Invoice 1027 Bill\'s Windsurf Shop Accounts Receivable (A/R) 85.00'
result = parse_journal_line(line)
print(result)
