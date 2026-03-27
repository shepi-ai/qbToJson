#!/usr/bin/env python3
"""Debug preprocessing to see why false accounts still appear"""
import pdfplumber
from generalLedgerConverter import GeneralLedgerConverter

# Get raw and preprocessed lines
converter = GeneralLedgerConverter()

with pdfplumber.open('sampleReports/2023GL/GeneralLedger-01.23.pdf') as pdf:
    all_text = ""
    for page in pdf.pages:
        text = page.extract_text()
        if text:
            all_text += text + "\n"

    raw_lines = all_text.split('\n')
    preprocessed_lines = converter._preprocess_pdf_lines(raw_lines)

    # Find "Workers Compensation" in preprocessed lines
    print("=== Finding 'Workers Compensation' in preprocessed lines ===\n")
    for i, line in enumerate(preprocessed_lines):
        if 'Workers Compensation' in line:
            print(f'Line {i}: {line}')
            if i + 1 < len(preprocessed_lines):
                print(f'Line {i+1}: {preprocessed_lines[i+1]}')
            if i + 2 < len(preprocessed_lines):
                print(f'Line {i+2}: {preprocessed_lines[i+2]}')
            print()

    # Find "Accounting" in preprocessed lines
    print("=== Finding 'Accounting' in preprocessed lines ===\n")
    for i, line in enumerate(preprocessed_lines):
        if line.strip() == 'Accounting':
            print(f'Line {i}: "{line}"')
            if i + 1 < len(preprocessed_lines):
                print(f'Line {i+1}: "{preprocessed_lines[i+1]}"')
            if i + 2 < len(preprocessed_lines):
                print(f'Line {i+2}: "{preprocessed_lines[i+2]}"')
            print()
