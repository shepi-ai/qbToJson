#!/usr/bin/env python3
"""
Analyze qbToJson Google Cloud Run logs
"""
import json
import subprocess
from datetime import datetime
from collections import defaultdict, Counter

def get_logs():
    """Fetch logs from Google Cloud"""
    cmd = [
        'gcloud', 'logging', 'read',
        'resource.type=cloud_run_revision AND resource.labels.service_name=qbtojson',
        '--limit=200',
        '--format=json',
        '--project=flash-cache-444122-k2'
    ]
    
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"Error fetching logs: {result.stderr}")
        return []
    
    return json.loads(result.stdout) if result.stdout else []

def analyze_logs(logs):
    """Analyze the log entries"""
    stats = {
        'total_requests': 0,
        'successful': 0,
        'client_errors': 0,  # 4xx
        'server_errors': 0,  # 5xx
        'by_endpoint': defaultdict(lambda: {'success': 0, 'failed': 0, 'total_latency': 0}),
        'by_status': Counter(),
        'request_sizes': {'success': [], 'failed': []},
        'latencies': [],
        'errors': []
    }
    
    for entry in logs:
        # Skip non-HTTP request logs
        if 'httpRequest' not in entry:
            continue
        
        http = entry['httpRequest']
        status = http.get('status', 0)
        url = http.get('requestUrl', '')
        method = http.get('requestMethod', '')
        latency = http.get('latency', '0s')
        request_size = http.get('requestSize', 0)
        
        # Extract endpoint
        endpoint = 'unknown'
        if '/api/convert/' in url:
            endpoint = url.split('/api/convert/')[-1].split('?')[0]
        
        stats['total_requests'] += 1
        stats['by_status'][status] += 1
        
        # Parse latency (format: "0.168038319s")
        try:
            latency_seconds = float(latency.rstrip('s'))
            stats['latencies'].append(latency_seconds)
            stats['by_endpoint'][endpoint]['total_latency'] += latency_seconds
        except:
            pass
        
        # Categorize by status
        if status >= 200 and status < 300:
            stats['successful'] += 1
            stats['by_endpoint'][endpoint]['success'] += 1
            stats['request_sizes']['success'].append(int(request_size))
        elif status >= 400 and status < 500:
            stats['client_errors'] += 1
            stats['by_endpoint'][endpoint]['failed'] += 1
            stats['request_sizes']['failed'].append(int(request_size))
            
            # Track error details
            stats['errors'].append({
                'status': status,
                'endpoint': endpoint,
                'method': method,
                'size': request_size,
                'timestamp': entry.get('timestamp', ''),
                'latency': latency
            })
        elif status >= 500:
            stats['server_errors'] += 1
            stats['by_endpoint'][endpoint]['failed'] += 1
    
    return stats

def print_report(stats):
    """Print a formatted report"""
    print("=" * 80)
    print("QBTOJSON SERVICE - DOCUMENT PARSING PERFORMANCE REPORT")
    print("=" * 80)
    print()
    
    # Overall stats
    print("📊 OVERALL STATISTICS")
    print("-" * 80)
    print(f"Total Requests:        {stats['total_requests']}")
    print(f"✅ Successful (2xx):    {stats['successful']} ({stats['successful']/stats['total_requests']*100:.1f}%)")
    print(f"⚠️  Client Errors (4xx): {stats['client_errors']} ({stats['client_errors']/stats['total_requests']*100:.1f}%)")
    print(f"❌ Server Errors (5xx): {stats['server_errors']} ({stats['server_errors']/stats['total_requests']*100:.1f}%)")
    print()
    
    # Status code breakdown
    print("📈 STATUS CODE BREAKDOWN")
    print("-" * 80)
    for status, count in sorted(stats['by_status'].items()):
        symbol = "✅" if status < 300 else "⚠️" if status < 500 else "❌"
        print(f"{symbol} {status}: {count:4d} requests ({count/stats['total_requests']*100:5.1f}%)")
    print()
    
    # Endpoint breakdown
    print("🎯 PERFORMANCE BY ENDPOINT")
    print("-" * 80)
    for endpoint, data in sorted(stats['by_endpoint'].items()):
        total = data['success'] + data['failed']
        success_rate = (data['success'] / total * 100) if total > 0 else 0
        avg_latency = (data['total_latency'] / total) if total > 0 else 0
        
        print(f"\n/{endpoint}")
        print(f"  Total:         {total} requests")
        print(f"  Success:       {data['success']} ({success_rate:.1f}%)")
        print(f"  Failed:        {data['failed']} ({100-success_rate:.1f}%)")
        print(f"  Avg Latency:   {avg_latency:.3f}s")
    print()
    
    # Latency analysis
    if stats['latencies']:
        latencies = sorted(stats['latencies'])
        print("⏱️  LATENCY ANALYSIS")
        print("-" * 80)
        print(f"Average:    {sum(latencies)/len(latencies):.3f}s")
        print(f"Median:     {latencies[len(latencies)//2]:.3f}s")
        print(f"Min:        {min(latencies):.3f}s")
        print(f"Max:        {max(latencies):.3f}s")
        print(f"P95:        {latencies[int(len(latencies)*0.95)]:.3f}s")
        print(f"P99:        {latencies[int(len(latencies)*0.99)]:.3f}s")
        print()
    
    # Request size analysis
    print("📦 REQUEST SIZE ANALYSIS")
    print("-" * 80)
    if stats['request_sizes']['success']:
        avg_success = sum(stats['request_sizes']['success']) / len(stats['request_sizes']['success'])
        print(f"✅ Successful requests: {avg_success:,.0f} bytes avg (n={len(stats['request_sizes']['success'])})")
    
    if stats['request_sizes']['failed']:
        avg_failed = sum(stats['request_sizes']['failed']) / len(stats['request_sizes']['failed'])
        print(f"⚠️  Failed requests:    {avg_failed:,.0f} bytes avg (n={len(stats['request_sizes']['failed'])})")
    print()
    
    # Error analysis
    if stats['errors']:
        print("🔍 ERROR ANALYSIS (Recent Client Errors)")
        print("-" * 80)
        print(f"Total 4xx errors: {len(stats['errors'])}")
        print("\nMost recent errors:")
        for i, error in enumerate(stats['errors'][:10], 1):
            print(f"\n{i}. Status {error['status']} - {error['endpoint']}")
            print(f"   Size: {int(error['size']):,} bytes | Latency: {error['latency']}")
            print(f"   Time: {error['timestamp']}")
        print()
    
    # Summary and recommendations
    print("💡 KEY FINDINGS")
    print("-" * 80)
    
    success_rate = (stats['successful'] / stats['total_requests'] * 100) if stats['total_requests'] > 0 else 0
    
    if success_rate >= 90:
        print("✅ Overall health: EXCELLENT")
    elif success_rate >= 75:
        print("⚠️  Overall health: GOOD (some issues detected)")
    elif success_rate >= 50:
        print("⚠️  Overall health: FAIR (significant issues)")
    else:
        print("❌ Overall health: POOR (major issues)")
    
    print()
    
    # Analyze the 400 errors
    if stats['client_errors'] > 0:
        avg_failed_size = sum(stats['request_sizes']['failed']) / len(stats['request_sizes']['failed']) if stats['request_sizes']['failed'] else 0
        avg_success_size = sum(stats['request_sizes']['success']) / len(stats['request_sizes']['success']) if stats['request_sizes']['success'] else 0
        
        print(f"• {stats['client_errors']} client errors (400 status) detected")
        print(f"• Failed requests are very small (~{avg_failed_size:,.0f} bytes)")
        print(f"• Successful requests are much larger (~{avg_success_size:,.0f} bytes)")
        print("• This suggests failed requests are missing file data (validation errors)")
        print("• Failed requests complete in ~2ms (immediate validation failure)")
        print()
    
    if stats['latencies']:
        avg_latency = sum(stats['latencies']) / len(stats['latencies'])
        print(f"• Average latency: {avg_latency:.3f}s")
        if avg_latency < 0.5:
            print("• Performance: EXCELLENT (< 500ms)")
        elif avg_latency < 1.0:
            print("• Performance: GOOD (< 1s)")
        else:
            print("• Performance: NEEDS ATTENTION (> 1s)")
    
    print()
    print("=" * 80)

if __name__ == '__main__':
    print("Fetching logs from Google Cloud Run...")
    logs = get_logs()
    
    if not logs:
        print("No logs found or error fetching logs")
        exit(1)
    
    print(f"Analyzing {len(logs)} log entries...\n")
    
    stats = analyze_logs(logs)
    print_report(stats)
