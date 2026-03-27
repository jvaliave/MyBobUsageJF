#!/usr/bin/env python3
"""
Analyze Aha! Epic data from Excel file and generate comprehensive PDF report with:
1. Epic status distribution
2. Epic tags (Squads) with count table and pie chart
3. Company Association line chart
"""

import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import matplotlib.patches as mpatches
from datetime import datetime
import requests
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

def download_aha_file():
    """Download the Aha! Excel file."""
    url = "https://bigblue.aha.io/products/PSAGE/bookmarks/file_exports/7621190615980411603.xlsx"
    filename = "aha_epics_export.xlsx"
    
    print(f"Downloading Aha! export from: {url}")
    
    # Get Aha! credentials from environment
    aha_token = os.getenv('AHA_API_KEY') or os.getenv('AHA_TOKEN')
    
    headers = {}
    if aha_token:
        headers['Authorization'] = f'Bearer {aha_token}'
    
    try:
        response = requests.get(url, headers=headers, allow_redirects=True)
        response.raise_for_status()
        
        with open(filename, 'wb') as f:
            f.write(response.content)
        
        print(f"✓ File downloaded: {filename}")
        return filename
    except Exception as e:
        print(f"Error downloading file: {e}")
        print("Checking if file already exists locally...")
        if os.path.exists(filename):
            print(f"✓ Using existing file: {filename}")
            return filename
        raise

def analyze_epic_status(df):
    """Analyze Epic status distribution."""
    # Find the status column (case-insensitive)
    status_col = None
    for col in df.columns:
        if 'status' in col.lower() and 'epic' in col.lower():
            status_col = col
            break
    
    if status_col and status_col in df.columns:
        status_counts = df[status_col].value_counts()
        print("\nEpic Status Distribution:")
        for status, count in status_counts.items():
            print(f"  {status}: {count}")
        return status_counts
    else:
        print(f"Warning: Status column not found. Available columns: {list(df.columns)}")
        return pd.Series()

def analyze_epic_tags(df):
    """Analyze Epic tags (Squads) and count."""
    # Find the tags column (case-insensitive)
    tag_col = None
    for col in df.columns:
        if 'tag' in col.lower() and 'epic' in col.lower():
            tag_col = col
            break
    
    if not tag_col:
        print(f"Warning: Tags column not found. Available columns: {list(df.columns)}")
        return pd.DataFrame(), pd.Series()
    
    # Parse tags (they might be comma-separated)
    all_tags = []
    epic_refs = []
    epic_urls = []
    
    # Find reference and URL columns (case-insensitive)
    ref_col = None
    url_col = None
    for col in df.columns:
        if 'reference' in col.lower() and 'epic' in col.lower():
            ref_col = col
        if 'url' in col.lower() and 'epic' in col.lower():
            url_col = col
    
    for idx, row in df.iterrows():
        tags = str(row[tag_col]) if pd.notna(row[tag_col]) else ""
        if tags and tags != 'nan':
            # Split by comma if multiple tags
            tag_list = [t.strip() for t in tags.split(',')]
            for tag in tag_list:
                if tag:
                    all_tags.append(tag)
                    # Get epic reference and URL
                    epic_ref = row.get(ref_col, '') if ref_col else ''
                    epic_url = row.get(url_col, '') if url_col else ''
                    epic_refs.append(epic_ref)
                    epic_urls.append(epic_url)
    
    # Create DataFrame
    tag_df = pd.DataFrame({
        'Squad': all_tags,
        'Epic Reference': epic_refs,
        'Epic URL': epic_urls
    })
    
    # Count by squad
    squad_counts = tag_df['Squad'].value_counts()
    
    print("\nEpic Tags (Squads) Distribution:")
    for squad, count in squad_counts.items():
        print(f"  {squad}: {count}")
    
    return tag_df, squad_counts

def analyze_company_association(df):
    """Analyze Company Association."""
    # Find company association column (case-insensitive)
    company_col = None
    for col in df.columns:
        if 'company' in col.lower() and 'association' in col.lower():
            company_col = col
            break
    
    if not company_col:
        print(f"Warning: Company Association column not found.")
        return pd.DataFrame()
    
    # Filter rows with company association
    company_df = df[df[company_col].notna()].copy()
    
    # Find reference and URL columns (case-insensitive)
    ref_col = None
    url_col = None
    for col in df.columns:
        if 'reference' in col.lower() and 'epic' in col.lower():
            ref_col = col
        if 'url' in col.lower() and 'epic' in col.lower():
            url_col = col
    
    # Get epic reference and URL
    result_df = pd.DataFrame()
    if ref_col:
        result_df['Epic Reference'] = company_df[ref_col]
    if url_col:
        result_df['Epic URL'] = company_df[url_col]
    
    result_df['Company Association'] = company_df[company_col]
    
    print(f"\nEpics with Company Association: {len(result_df)}")
    
    return result_df

def create_pdf_report(df, status_counts, tag_df, squad_counts, company_df):
    """Create comprehensive PDF report."""
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    pdf_filename = f'lineage_epic_analysis_{timestamp}.pdf'
    
    with PdfPages(pdf_filename) as pdf:
        # Title Page
        fig_title = plt.figure(figsize=(11, 8.5))
        ax_title = fig_title.add_subplot(111)
        ax_title.axis('off')
        
        # Main title
        report_date = datetime.now().strftime('%B %d, %Y')
        title_text = f"Lineage 2026 Epic Planning & Execution Report"
        date_text = f"Generated on: {report_date}"
        
        ax_title.text(0.5, 0.6, title_text,
                     ha='center', va='center', fontsize=24, fontweight='bold',
                     transform=ax_title.transAxes)
        
        # Add underline
        ax_title.plot([0.15, 0.85], [0.57, 0.57], 'k-', linewidth=2, transform=ax_title.transAxes)
        
        ax_title.text(0.5, 0.5, date_text,
                     ha='center', va='center', fontsize=14,
                     transform=ax_title.transAxes)
        
        # Summary stats
        summary_text = f"Total Epics Analyzed: {len(df)}\n"
        summary_text += f"Epic Status Categories: {len(status_counts)}\n"
        summary_text += f"Squad Tags: {len(squad_counts)}\n"
        summary_text += f"Epics with Company Association: {len(company_df)}"
        
        ax_title.text(0.5, 0.35, summary_text,
                     ha='center', va='center', fontsize=12,
                     transform=ax_title.transAxes,
                     bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.3))
        
        pdf.savefig(fig_title, bbox_inches='tight')
        plt.close()
        
        # Page 1: Epic Status Distribution
        if not status_counts.empty:
            fig = plt.figure(figsize=(11, 8.5))
            ax = fig.add_subplot(111)
            
            # Create bar chart
            bars = ax.bar(range(len(status_counts)), status_counts.values, 
                         color='steelblue', edgecolor='black', linewidth=1.5)
            
            ax.set_xlabel('Epic Status', fontsize=12, fontweight='bold')
            ax.set_ylabel('Count', fontsize=12, fontweight='bold')
            ax.set_title('Epic Status Distribution', fontsize=14, fontweight='bold', pad=20)
            ax.set_xticks(range(len(status_counts)))
            ax.set_xticklabels(status_counts.index, rotation=45, ha='right')
            
            # Add value labels
            for bar, value in zip(bars, status_counts.values):
                ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5,
                       str(value), ha='center', va='bottom', fontweight='bold')
            
            ax.grid(axis='y', alpha=0.3, linestyle='--')
            plt.tight_layout()
            pdf.savefig(fig, bbox_inches='tight')
            plt.close()
        
        # Page 2: Epic Tags (Squads) - 3D-style Pie Chart
        if not squad_counts.empty:
            fig = plt.figure(figsize=(12, 10))
            ax1 = fig.add_subplot(111)
            
            # Generate colors
            colors = plt.cm.Set3(range(len(squad_counts)))
            
            # Create explode effect for 3D appearance (slightly separate largest slice)
            explode = [0.05 if i == 0 else 0 for i in range(len(squad_counts))]
            
            # Create pie chart with 3D-style effects
            wedges, texts, autotexts = ax1.pie(squad_counts.values,
                                               labels=squad_counts.index,
                                               autopct='%1.1f%%',
                                               colors=colors,
                                               startangle=90,
                                               explode=explode,
                                               shadow=True,
                                               textprops={'fontsize': 9})
            
            # Style the text for better visibility
            for text in texts:
                text.set_fontsize(9)
                text.set_fontweight('bold')
            for autotext in autotexts:
                autotext.set_color('white')
                autotext.set_fontweight('bold')
                autotext.set_fontsize(8)
            
            ax1.set_title('Epic Distribution by Squad (3D View)', fontsize=16, fontweight='bold', pad=20)
            
            plt.tight_layout()
            pdf.savefig(fig, bbox_inches='tight')
            plt.close()
            
            # Page 3: Epic Count Table
            fig = plt.figure(figsize=(11, 8.5))
            ax2 = fig.add_subplot(111)
            ax2.axis('off')
            
            # Create table data
            table_data = [['Squad', 'Epic Count']]
            for squad, count in squad_counts.items():
                table_data.append([squad, str(count)])
            
            table = ax2.table(cellText=table_data, cellLoc='left',
                            loc='center', colWidths=[0.7, 0.3])
            table.auto_set_font_size(False)
            table.set_fontsize(10)
            table.scale(1, 2)
            
            # Style header row
            for i in range(2):
                table[(0, i)].set_facecolor('#366092')
                table[(0, i)].set_text_props(weight='bold', color='white')
            
            # Alternate row colors
            for i in range(1, len(table_data)):
                for j in range(2):
                    if i % 2 == 0:
                        table[(i, j)].set_facecolor('#f0f0f0')
            
            ax2.set_title('Epic Count by Squad', fontsize=16, fontweight='bold', pad=20)
            
            plt.tight_layout()
            pdf.savefig(fig, bbox_inches='tight')
            plt.close()
        
        # Page 4+: Company Association Bar Chart and List
        if not company_df.empty:
            # Bar chart showing count by company
            fig = plt.figure(figsize=(14, 10))
            ax = fig.add_subplot(111)
            
            # Count epics by company
            company_counts = company_df['Company Association'].value_counts().sort_values(ascending=False)
            
            # Create bar chart
            bars = ax.bar(range(len(company_counts)), company_counts.values,
                         color='steelblue', edgecolor='black', linewidth=1.5)
            
            ax.set_xlabel('Company Association', fontsize=12, fontweight='bold')
            ax.set_ylabel('Count of Epics', fontsize=12, fontweight='bold')
            ax.set_title(f'Epics by Company Association (Total: {len(company_df)} epics)',
                        fontsize=14, fontweight='bold', pad=20)
            
            # Set x-axis labels
            ax.set_xticks(range(len(company_counts)))
            ax.set_xticklabels(company_counts.index, rotation=45, ha='right', fontsize=10)
            
            # Add value labels on bars
            for bar, value in zip(bars, company_counts.values):
                ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5,
                       str(value), ha='center', va='bottom', fontweight='bold', fontsize=10)
            
            ax.grid(axis='y', alpha=0.3, linestyle='--')
            plt.tight_layout()
            pdf.savefig(fig, bbox_inches='tight')
            plt.close()
            
            # List of epics with company association
            fig = plt.figure(figsize=(11, 8.5))
            ax = fig.add_subplot(111)
            ax.axis('off')
            
            ax.text(0.5, 0.95, 'Epics with Company Association', 
                   ha='center', va='top', fontsize=16, fontweight='bold',
                   transform=ax.transAxes)
            
            y_position = 0.88
            line_height = 0.04
            
            for idx, row in company_df.iterrows():
                epic_ref = row.get('Epic Reference', 'N/A')
                epic_url = row.get('Epic URL', 'N/A')
                company = row.get('Company Association', 'N/A')
                
                # Epic reference
                ax.text(0.05, y_position, f"Epic: {epic_ref}",
                       ha='left', va='top', fontsize=10, fontweight='bold',
                       transform=ax.transAxes)
                y_position -= line_height
                
                # URL
                ax.text(0.08, y_position, f"URL: {epic_url}",
                       ha='left', va='top', fontsize=8, color='blue',
                       transform=ax.transAxes, style='italic')
                y_position -= line_height
                
                # Company
                ax.text(0.08, y_position, f"Company: {company}",
                       ha='left', va='top', fontsize=8,
                       transform=ax.transAxes)
                y_position -= line_height * 1.5
                
                # Check if we need a new page
                if y_position < 0.1:
                    pdf.savefig(fig, bbox_inches='tight')
                    plt.close()
                    
                    fig = plt.figure(figsize=(11, 8.5))
                    ax = fig.add_subplot(111)
                    ax.axis('off')
                    ax.text(0.5, 0.95, 'Epics with Company Association (continued)', 
                           ha='center', va='top', fontsize=16, fontweight='bold',
                           transform=ax.transAxes)
                    y_position = 0.88
            
            pdf.savefig(fig, bbox_inches='tight')
            plt.close()
    
    print(f"\n✓ PDF report saved: {pdf_filename}")
    return pdf_filename

def main():
    """Main function."""
    print("=" * 70)
    print("Aha! Epic Analysis Report Generator")
    print("=" * 70)
    
    try:
        # Use the downloaded Aha! export file
        filename = "aha_list_features_260326054547.xlsx"
        
        if not os.path.exists(filename):
            print(f"File not found: {filename}")
            print("Please download the file and place it in the current directory.")
            return
        
        # Read Excel file
        print(f"\nReading Excel file: {filename}")
        df = pd.read_excel(filename)
        print(f"Total records: {len(df)}")
        print(f"Columns: {list(df.columns)}")
        
        # Perform analyses
        status_counts = analyze_epic_status(df)
        tag_df, squad_counts = analyze_epic_tags(df)
        company_df = analyze_company_association(df)
        
        # Create PDF report
        pdf_file = create_pdf_report(df, status_counts, tag_df, squad_counts, company_df)
        
        print("\n" + "=" * 70)
        print("Analysis complete!")
        print(f"Report saved: {pdf_file}")
        print("=" * 70)
        
    except Exception as e:
        print(f"\nError: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()

# Made with Bob
