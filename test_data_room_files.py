#!/usr/bin/env python3
"""
Test QuickBooks files from the data room against local Docker container.
This script processes all supported QB files and generates a comprehensive report.
"""

import os
import requests
import json
from pathlib import Path
from datetime import datetime
import sys

# Configuration
BASE_URL = "http://localhost:8080"
API_KEY = "test-local-key"
DATA_ROOM = "/Users/araboin/Documents/PotomacHeritage/Artistic Kitchen and Bath/Artistic Kitchen & Bath /QofE Data Room"

# Folder to endpoint mapping
FOLDER_MAPPINGS = {
    "7 - P&L by Month": "profit-loss",
    "9 - Balance Sheet by Month": "balance-sheet",
    "10 - Balance Sheet Detail": "balance-sheet",
    "11 - General Ledger": "general-ledger",
    "12 - Cash Flow": "cash-flow",
    "22 - AR Aging": "accounts-receivable",
    "25 - AP Aging": "accounts-payable",
    "27 - Sales by Customer": "customer-concentration",
}

# Results storage
results = {
    "success": [],
    "failed": [],
    "skipped": [],
    "total_files": 0,
    "start_time": datetime.now().isoformat()
}

def test_health():
    """Check if the Docker container is running."""
    try:
        response = requests.get(f"{BASE_URL}/health", timeout=5)
        if response.status_code == 200:
            print("✅ Docker container is running and healthy\n")
            return True
        else:
            print(f"❌ Docker container returned status {response.status_code}\n")
            return False
    except requests.exceptions.RequestException as e:
        print(f"❌ Cannot connect to Docker container at {BASE_URL}")
        print(f"   Error: {e}")
        print("\n   Make sure the container is running:")
        print("   docker ps | grep qbtojson-local\n")
        return False

def get_endpoint_for_file(file_path):
    """Determine the API endpoint based on file location."""
    for folder, endpoint in FOLDER_MAPPINGS.items():
        if folder in str(file_path):
            return endpoint
    return None

def find_quickbooks_files():
    """Scan data room for QuickBooks files."""
    files = []
    data_room_path = Path(DATA_ROOM)
    
    if not data_room_path.exists():
        print(f"❌ Data room path not found: {DATA_ROOM}")
        return files
    
    # Find all .xlsx files in mapped folders
    for folder in FOLDER_MAPPINGS.keys():
        folder_path = data_room_path / folder
        if folder_path.exists():
            xlsx_files = list(folder_path.glob("*.xlsx"))
            for file in xlsx_files:
                endpoint = get_endpoint_for_file(file)
                if endpoint:
                    files.append({
                        "path": str(file),
                        "name": file.name,
                        "folder": folder,
                        "endpoint": endpoint
                    })
    
    return files

def test_file(file_info):
    """Test a single file against the API."""
    file_path = file_info["path"]
    endpoint = file_info["endpoint"]
    file_name = file_info["name"]
    
    print(f"Testing: {file_name}")
    print(f"  Endpoint: /api/convert/{endpoint}")
    
    try:
        with open(file_path, 'rb') as f:
            files = {'file': (file_name, f, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
            headers = {'x-api-key': API_KEY}
            
            response = requests.post(
                f"{BASE_URL}/api/convert/{endpoint}",
                files=files,
                headers=headers,
                timeout=60
            )
            
            if response.status_code == 200:
                data = response.json()
                if data.get('success'):
                    # Extract meaningful info from response
                    record_count = 0
                    if 'count' in data:
                        record_count = data['count']
                    elif 'months' in data:
                        record_count = data['months']
                    elif 'data' in data:
                        if isinstance(data['data'], list):
                            record_count = len(data['data'])
                        elif isinstance(data['data'], dict) and 'monthlyReports' in data['data']:
                            record_count = len(data['data']['monthlyReports'])
                    
                    print(f"  ✅ SUCCESS - {record_count} records/months\n")
                    results["success"].append({
                        "file": file_name,
                        "folder": file_info["folder"],
                        "endpoint": endpoint,
                        "record_count": record_count
                    })
                    return True
                else:
                    error = data.get('error', 'Unknown error')
                    print(f"  ❌ FAILED - {error}\n")
                    results["failed"].append({
                        "file": file_name,
                        "folder": file_info["folder"],
                        "endpoint": endpoint,
                        "error": error,
                        "details": data.get('details', '')
                    })
                    return False
            else:
                error_text = response.text[:200]
                print(f"  ❌ FAILED - HTTP {response.status_code}")
                print(f"     {error_text}\n")
                results["failed"].append({
                    "file": file_name,
                    "folder": file_info["folder"],
                    "endpoint": endpoint,
                    "error": f"HTTP {response.status_code}",
                    "details": error_text
                })
                return False
                
    except Exception as e:
        print(f"  ❌ EXCEPTION - {str(e)}\n")
        results["failed"].append({
            "file": file_name,
            "folder": file_info["folder"],
            "endpoint": endpoint,
            "error": "Exception",
            "details": str(e)
        })
        return False

def generate_report():
    """Generate and display test results report."""
    results["end_time"] = datetime.now().isoformat()
    
    print("\n" + "="*80)
    print("TEST RESULTS SUMMARY")
    print("="*80)
    
    total = results["total_files"]
    success_count = len(results["success"])
    failed_count = len(results["failed"])
    skipped_count = len(results["skipped"])
    
    print(f"\nTotal Files Tested: {total}")
    print(f"✅ Successful: {success_count} ({success_count/total*100:.1f}%)" if total > 0 else "✅ Successful: 0")
    print(f"❌ Failed: {failed_count} ({failed_count/total*100:.1f}%)" if total > 0 else "❌ Failed: 0")
    print(f"⏭️  Skipped: {skipped_count}")
    
    if results["success"]:
        print("\n" + "-"*80)
        print("✅ SUCCESSFUL FILES")
        print("-"*80)
        for item in results["success"]:
            print(f"  • {item['file']}")
            print(f"    Endpoint: {item['endpoint']}")
            print(f"    Records: {item['record_count']}")
            print()
    
    if results["failed"]:
        print("\n" + "-"*80)
        print("❌ FAILED FILES")
        print("-"*80)
        for item in results["failed"]:
            print(f"  • {item['file']}")
            print(f"    Endpoint: {item['endpoint']}")
            print(f"    Error: {item['error']}")
            if item['details']:
                print(f"    Details: {item['details'][:100]}")
            print()
    
    # Save detailed report to file
    report_file = "test_results.json"
    with open(report_file, 'w') as f:
        json.dump(results, f, indent=2)
    
    print(f"\n📄 Detailed results saved to: {report_file}")
    print("="*80 + "\n")

def main():
    """Main execution function."""
    print("\n" + "="*80)
    print("QUICKBOOKS DATA ROOM FILE TESTING")
    print("="*80)
    print(f"Data Room: {DATA_ROOM}")
    print(f"API Base URL: {BASE_URL}")
    print("="*80 + "\n")
    
    # Check if container is running
    if not test_health():
        sys.exit(1)
    
    # Find all QuickBooks files
    print("Scanning data room for QuickBooks files...")
    files = find_quickbooks_files()
    
    if not files:
        print("❌ No QuickBooks files found in data room")
        sys.exit(1)
    
    results["total_files"] = len(files)
    
    print(f"Found {len(files)} QuickBooks files to test\n")
    print("="*80 + "\n")
    
    # Test each file
    for i, file_info in enumerate(files, 1):
        print(f"[{i}/{len(files)}] ", end="")
        test_file(file_info)
    
    # Generate report
    generate_report()

if __name__ == "__main__":
    main()
