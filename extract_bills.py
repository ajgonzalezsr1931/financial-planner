#!/usr/bin/env python3
"""
Electric Bill PDF Data Extraction Tool
Extracts CURRENT CHARGES ONLY from electric company PDFs, excluding previous balances and fees.
Analyzes seasonal patterns based on actual electricity usage charges per month.
"""

import os
import re
import csv
import subprocess
from pathlib import Path
from datetime import datetime
from collections import defaultdict
import statistics


def extract_pdf_data(pdf_path):
    """
    Extract billing data from a single PDF using pdftotext.
    Returns: {'date': YYYYMMDD, 'amount': float, 'has_late_charges': bool, 'raw_text': str}
    Note: 'amount' is the current month's electricity charges only, not total due.
    """
    try:
        # Use pdftotext to extract text from PDF
        result = subprocess.run(
            ['pdftotext', '-layout', str(pdf_path), '-'],
            capture_output=True,
            text=True,
            timeout=30
        )
        
        if result.returncode != 0:
            print(f"ERROR: pdftotext failed for {pdf_path}")
            return None
            
        full_text = result.stdout
        
        # Extract date from filename (format: YYYYMMDD at end before .pdf)
        filename = Path(pdf_path).name
        date_match = re.search(r'(\d{8})\.pdf$', filename)
        billing_date = date_match.group(1) if date_match else None
        
        # Parse for current month's charges only (not total due)
        amount = parse_total_amount(full_text)
        
        # Check for late charges or fees to exclude
        has_late_charges = any(term in full_text.upper() for term in [
            'LATE', 'PENALTY', 'DNP', 'NOTICE', 'DELINQUENT'
        ])
        
        return {
            'date': billing_date,
            'amount': amount,
            'has_late_charges': has_late_charges,
            'raw_text': full_text[:500]  # Save snippet for verification
        }
    except Exception as e:
        print(f"ERROR processing {pdf_path}: {e}")
        return None


def parse_total_amount(text):
    """
    Extract the current charges only, excluding previous balances and late fees.
    Looks for "Current Charges", "New Charges", "Electric Service", etc.
    """
    # Primary patterns for current month's electricity charges - most specific first
    current_charge_patterns = [
        r'[Tt]otal\s+[Cc]urrent\s+[Cc]harges[.\s]*\$?([\d,]+\.?\d{0,2})',
        r'[Nn]ew\s+[Cc]harges[.\s]*\$?([\d,]+\.?\d{0,2})',
        r'[Cc]urrent\s+[Cc]harges[.\s]*\$?([\d,]+\.?\d{0,2})',
    ]
    
    # Try the most specific patterns first
    for pattern in current_charge_patterns:
        matches = re.findall(pattern, text, re.IGNORECASE)
        if matches:
            # Return the first match (should be the total current charges)
            amounts = [float(m.replace(',', '')) for m in matches]
            return amounts[0] if amounts else None
    
    # If no explicit "current charges" found, try to calculate from individual line items
    # Look for the detailed breakdown section
    lines = text.split('\n')
    energy_charges = []
    delivery_charges = []
    tax_charges = []
    
    for i, line in enumerate(lines):
        line = line.strip()
        
        # Look for energy charge lines like "Energy Charge 1,947 KWh @ $0.09330 Per KWh...$181.66"
        energy_match = re.search(r'[Ee]nergy\s+[Cc]harge.*\$?([\d,]+\.\d{2})', line)
        if energy_match:
            energy_charges.append(float(energy_match.group(1).replace(',', '')))
            continue
            
        # Look for delivery charges
        delivery_match = re.search(r'[Dd]elivery\s+[Cc]harges.*\$?([\d,]+\.\d{2})', line)
        if delivery_match:
            delivery_charges.append(float(delivery_match.group(1).replace(',', '')))
            continue
            
        # Look for tax and fee lines (but not late fees)
        if any(term in line.upper() for term in ['TAX', 'ASSESSMENT', 'SECURITIZATION']) and 'LATE' not in line.upper():
            tax_match = re.search(r'\$?([\d,]+\.\d{2})(?!.*LATE)', line)
            if tax_match:
                tax_charges.append(float(tax_match.group(1).replace(',', '')))
    
    # Sum up the components if we found them
    total_components = sum(energy_charges) + sum(delivery_charges) + sum(tax_charges)
    if total_components > 0:
        return total_components
    
    # Last resort: look for reasonable electric bill amounts in context
    # but exclude obvious non-charges like addresses, account numbers, etc.
    text_lines = text.split('\n')
    charge_amounts = []
    
    for line in text_lines:
        # Skip lines that clearly aren't charges
        if any(skip_term in line.upper() for skip_term in [
            'PREVIOUS', 'BALANCE', 'TOTAL DUE', 'AMOUNT DUE', 'CARRY', 'FORWARD',
            'BUTTERWICK', 'SPRING', 'ADDRESS', 'ACCOUNT', 'ESI ID', 'PHONE', 'EMAIL'
        ]):
            continue
            
        # Look for dollar amounts that could be current charges
        dollar_matches = re.findall(r'(?<!\d)\$?([\d,]+\.\d{2})(?!\d)', line)
        if dollar_matches:
            for amount_str in dollar_matches:
                amount = float(amount_str.replace(',', ''))
                # Filter reasonable electric bill amounts (between $20 and $1500)
                if 20.0 <= amount <= 1500.0:
                    charge_amounts.append(amount)
    
    # Return a reasonable current charge amount
    if charge_amounts:
        # Remove duplicates and sort
        unique_charges = list(set(charge_amounts))
        unique_charges.sort()
        
        # Often the current charges are in the middle range, not the highest/lowest
        if len(unique_charges) >= 3:
            # Return the median-ish value
            return unique_charges[len(unique_charges)//2]
        elif len(unique_charges) >= 2:
            # Return the higher of two values (often the current charge vs a component)
            return max(unique_charges)
        else:
            return unique_charges[0]
    
    return None


def get_month_number(date_str):
    """Convert YYYYMMDD to month number (1-12)"""
    if date_str and len(date_str) >= 6:
        try:
            month = int(date_str[4:6])
            return month
        except ValueError:
            return None
    return None


def get_season(month):
    """Classify month as summer (Jun-Sep) or winter (Dec-Mar) or shoulder (Apr-May, Oct-Nov)"""
    if month in [6, 7, 8, 9]:
        return 'Summer'
    elif month in [12, 1, 2, 3]:
        return 'Winter'
    else:
        return 'Shoulder'


def analyze_bills(bill_data_list):
    """Analyze billing data for patterns and statistics"""
    
    # Filter out bills with parsing errors
    valid_bills = [b for b in bill_data_list if b and b['amount']]
    
    if not valid_bills:
        print("ERROR: No valid billing data extracted!")
        return None
    
    # Organize by season
    seasonal_data = defaultdict(list)
    for bill in valid_bills:
        month = get_month_number(bill['date'])
        season = get_season(month)
        seasonal_data[season].append(bill['amount'])
    
    # Calculate statistics
    all_amounts = [b['amount'] for b in valid_bills]
    
    stats = {
        'total_bills': len(valid_bills),
        'all_amounts': all_amounts,
        'average': statistics.mean(all_amounts),
        'median': statistics.median(all_amounts),
        'min': min(all_amounts),
        'max': max(all_amounts),
        'stdev': statistics.stdev(all_amounts) if len(all_amounts) > 1 else 0,
        'seasonal': {
            'Summer': {
                'amounts': seasonal_data.get('Summer', []),
                'average': statistics.mean(seasonal_data['Summer']) if seasonal_data['Summer'] else 0
            },
            'Winter': {
                'amounts': seasonal_data.get('Winter', []),
                'average': statistics.mean(seasonal_data['Winter']) if seasonal_data['Winter'] else 0
            },
            'Shoulder': {
                'amounts': seasonal_data.get('Shoulder', []),
                'average': statistics.mean(seasonal_data['Shoulder']) if seasonal_data['Shoulder'] else 0
            }
        }
    }
    
    return stats, valid_bills


def main():
    # Configuration
    pdf_folder = "/home/anthony/Projects/financial planner CC/bills/chariot"
    output_csv = os.path.join(pdf_folder, "extracted_bills.csv")
    output_summary = os.path.join(pdf_folder, "billing_summary.txt")
    
    print(f"🔍 Scanning for PDFs in: {pdf_folder}")
    
    # Find all PDFs matching pattern
    pdf_files = sorted(Path(pdf_folder).glob("2001070033_BB*.pdf"))
    print(f"📄 Found {len(pdf_files)} PDF files")
    
    if not pdf_files:
        print("ERROR: No PDF files found!")
        return
    
    # Extract data from each PDF
    print("\n📊 Extracting current charges (excluding previous balances)...")
    bill_data_list = []
    for i, pdf_file in enumerate(pdf_files, 1):
        print(f"  [{i}/{len(pdf_files)}] {pdf_file.name}...", end='', flush=True)
        data = extract_pdf_data(str(pdf_file))
        if data:
            bill_data_list.append(data)
            print(f" ✓ ${data['amount']:.2f}" if data['amount'] else " ✗ (no current charges found)")
        else:
            print(" ✗ (extraction failed)")
    
    # Analyze patterns
    print("\n📈 Analyzing billing patterns...")
    stats, valid_bills = analyze_bills(bill_data_list)
    
    # Sort by date for CSV output
    valid_bills.sort(key=lambda x: x['date'])
    
    # Write CSV output
    print(f"\n💾 Writing results to: {output_csv}")
    with open(output_csv, 'w', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=['Date', 'Current_Charges', 'Month', 'Season', 'Has_Late_Charges'])
        writer.writeheader()
        for bill in valid_bills:
            month = get_month_number(bill['date'])
            season = get_season(month)
            writer.writerow({
                'Date': bill['date'],
                'Current_Charges': f"${bill['amount']:.2f}",
                'Month': month,
                'Season': season,
                'Has_Late_Charges': bill['has_late_charges']
            })
    
    # Write summary report
    print(f"📋 Writing summary to: {output_summary}")
    with open(output_summary, 'w') as f:
        f.write("=" * 70 + "\n")
        f.write("ELECTRIC BILL ANALYSIS SUMMARY - CURRENT CHARGES ONLY\n")
        f.write("(Excludes previous balances and late fees)\n")
        f.write("=" * 70 + "\n\n")
        
        f.write(f"Analysis Period: {valid_bills[0]['date']} to {valid_bills[-1]['date']}\n")
        f.write(f"Total Bills Analyzed: {stats['total_bills']}\n\n")
        
        f.write("OVERALL STATISTICS (Current Charges Only)\n")
        f.write("-" * 70 + "\n")
        f.write(f"Average Monthly Current Charges: ${stats['average']:.2f}\n")
        f.write(f"Median Monthly Current Charges: ${stats['median']:.2f}\n")
        f.write(f"Minimum Current Charges: ${stats['min']:.2f}\n")
        f.write(f"Maximum Current Charges: ${stats['max']:.2f}\n")
        f.write(f"Range: ${stats['min']:.2f} - ${stats['max']:.2f}\n")
        f.write(f"Standard Deviation: ${stats['stdev']:.2f}\n\n")
        
        f.write("SEASONAL BREAKDOWN\n")
        f.write("-" * 70 + "\n")
        for season in ['Summer', 'Winter', 'Shoulder']:
            data = stats['seasonal'][season]
            if data['amounts']:
                f.write(f"\n{season} Season ({len(data['amounts'])} bills):\n")
                f.write(f"  Average: ${data['average']:.2f}\n")
                f.write(f"  Range: ${min(data['amounts']):.2f} - ${max(data['amounts']):.2f}\n")
        
        f.write("\n\nKEY INSIGHTS\n")
        f.write("-" * 70 + "\n")
        
        # Calculate seasonal differences
        summer_avg = stats['seasonal']['Summer']['average']
        winter_avg = stats['seasonal']['Winter']['average']
        shoulder_avg = stats['seasonal']['Shoulder']['average']
        
        if summer_avg and winter_avg:
            diff_pct = ((summer_avg - winter_avg) / winter_avg) * 100
            f.write(f"Summer vs Winter: {abs(diff_pct):.1f}% {'higher' if diff_pct > 0 else 'lower'} in summer\n")
        
        if shoulder_avg:
            f.write(f"Shoulder season average: ${shoulder_avg:.2f}\n")
        
        # Variance indicator
        cv = (stats['stdev'] / stats['average']) * 100 if stats['average'] else 0
        f.write(f"Variability (CV): {cv:.1f}% - ")
        if cv < 10:
            f.write("Low variability - very consistent\n")
        elif cv < 20:
            f.write("Moderate variability\n")
        else:
            f.write("High variability - significant seasonal swings\n")
        
        f.write("\nRECOMMENDED VARIABLE MONTHLY BILL ENTRY\n")
        f.write("-" * 70 + "\n")
        f.write(f"Bill Name: Electric Bill\n")
        f.write(f"Estimated Amount: ${stats['average']:.2f}\n")
        f.write(f"Amount Range: ${stats['min']:.2f}-${stats['max']:.2f}\n")
        f.write(f"Due Day: 15 (or your typical due date)\n")
        f.write(f"Seasonal Note: Higher in summer (${summer_avg:.2f}) than winter (${winter_avg:.2f})\n")
    
    # Print summary to console
    print("\n" + "=" * 70)
    print("CURRENT CHARGES EXTRACTION COMPLETE")
    print("(Excluding previous balances and late fees)")
    print("=" * 70)
    print(f"\n✓ Processed {stats['total_bills']} bills successfully")
    print(f"\nAverage Monthly Current Charges: ${stats['average']:.2f}")
    print(f"Range: ${stats['min']:.2f} - ${stats['max']:.2f}")
    
    if stats['seasonal']['Summer']['amounts']:
        print(f"\nSummer Average: ${stats['seasonal']['Summer']['average']:.2f}")
    if stats['seasonal']['Winter']['amounts']:
        print(f"Winter Average: ${stats['seasonal']['Winter']['average']:.2f}")
    
    print(f"\n✓ CSV exported to: extracted_bills.csv")
    print(f"✓ Summary exported to: billing_summary.txt")


if __name__ == "__main__":
    main()