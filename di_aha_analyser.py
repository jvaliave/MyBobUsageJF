#!/usr/bin/env python3
"""
DI AHA ANALYSER - Master Script
Generates comprehensive epic analysis reports for:
1. IKC 2026 Epic Planning & Execution Report
2. Lineage 2026 Epic Planning & Execution Report

Includes blockers page if any epics have "Blocked" status.
"""

import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from datetime import datetime
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

def analyze_epic_status(df):
    """Analyze Epic status distribution."""
    status_col = None
    for col in df.columns:
        if 'status' in col.lower() and 'epic' in col.lower():
            status_col = col
            break
    
    if status_col and status_col in df.columns:
        status_counts = df[status_col].value_counts()
        return status_counts, status_col
    return pd.Series(), None

def analyze_epic_tags(df):
    """Analyze Epic tags (Squads) and count."""
    tag_col = None
    for col in df.columns:
        if 'tag' in col.lower() and 'epic' in col.lower():
            tag_col = col
            break
    
    if not tag_col:
        return pd.DataFrame(), pd.Series()
    
    ref_col = None
    url_col = None
    for col in df.columns:
        if 'reference' in col.lower() and 'epic' in col.lower():
            ref_col = col
        if 'url' in col.lower() and 'epic' in col.lower():
            url_col = col
    
    all_tags = []
    epic_refs = []
    epic_urls = []
    
    for idx, row in df.iterrows():
        tags = str(row[tag_col]) if pd.notna(row[tag_col]) else ""
        if tags and tags != 'nan':
            tag_list = [t.strip() for t in tags.split(',')]
            for tag in tag_list:
                if tag:
                    all_tags.append(tag)
                    epic_ref = row.get(ref_col, '') if ref_col else ''
                    epic_url = row.get(url_col, '') if url_col else ''
                    epic_refs.append(epic_ref)
                    epic_urls.append(epic_url)
    
    tag_df = pd.DataFrame({
        'Squad': all_tags,
        'Epic Reference': epic_refs,
        'Epic URL': epic_urls
    })
    
    squad_counts = tag_df['Squad'].value_counts()
    return tag_df, squad_counts

def analyze_company_association(df):
    """Analyze Company Association."""
    company_col = None
    for col in df.columns:
        if 'company' in col.lower() and 'association' in col.lower():
            company_col = col
            break
    
    if not company_col:
        return pd.DataFrame()
    
    company_df = df[df[company_col].notna()].copy()
    
    ref_col = None
    url_col = None
    for col in df.columns:
        if 'reference' in col.lower() and 'epic' in col.lower():
            ref_col = col
        if 'url' in col.lower() and 'epic' in col.lower():
            url_col = col
    
    result_df = pd.DataFrame()
    if ref_col:
        result_df['Epic Reference'] = company_df[ref_col]
    if url_col:
        result_df['Epic URL'] = company_df[url_col]
    
    result_df['Company Association'] = company_df[company_col]
    return result_df

def get_blocked_epics(df, status_col):
    """Get all epics with Blocked status."""
    if not status_col or status_col not in df.columns:
        return pd.DataFrame()
    
    blocked_df = df[df[status_col].str.lower() == 'blocked'].copy()
    
    if blocked_df.empty:
        return pd.DataFrame()
    
    # Find relevant columns
    ref_col = None
    name_col = None
    url_col = None
    git_col = None
    tag_col = None
    
    for col in df.columns:
        if 'reference' in col.lower() and 'epic' in col.lower():
            ref_col = col
        if 'name' in col.lower() and 'epic' in col.lower():
            name_col = col
        if 'url' in col.lower() and 'epic' in col.lower() and 'git' not in col.lower():
            url_col = col
        if 'git' in col.lower() and 'url' in col.lower():
            git_col = col
        if 'tag' in col.lower() and 'epic' in col.lower():
            tag_col = col
    
    result_df = pd.DataFrame()
    if ref_col:
        result_df['Epic Reference'] = blocked_df[ref_col]
    if name_col:
        result_df['Epic Name'] = blocked_df[name_col]
    if url_col:
        result_df['Epic URL'] = blocked_df[url_col]
    if git_col:
        result_df['Git URL'] = blocked_df[git_col]
    if tag_col:
        result_df['Epic Tags'] = blocked_df[tag_col]
    
    return result_df

def analyze_by_release(df):
    """Analyze epics grouped by release name with status color coding."""
    release_col = None
    ref_col = None
    name_col = None
    status_col = None
    tag_col = None
    url_col = None
    git_col = None
    
    for col in df.columns:
        if 'release' in col.lower() and 'name' in col.lower():
            release_col = col
        if 'reference' in col.lower() and 'epic' in col.lower():
            ref_col = col
        if 'name' in col.lower() and 'epic' in col.lower() and 'release' not in col.lower():
            name_col = col
        if 'status' in col.lower() and 'epic' in col.lower():
            status_col = col
        if 'tag' in col.lower() and 'epic' in col.lower():
            tag_col = col
        if 'url' in col.lower() and 'epic' in col.lower() and 'git' not in col.lower():
            url_col = col
        if 'git' in col.lower() and 'url' in col.lower():
            git_col = col
    
    if not release_col or release_col not in df.columns:
        return pd.DataFrame()
    
    # Group by release
    release_data = []
    for release in df[release_col].dropna().unique():
        release_epics = df[df[release_col] == release].copy()
        
        for idx, row in release_epics.iterrows():
            epic_ref = row.get(ref_col, 'N/A') if ref_col else 'N/A'
            epic_name = row.get(name_col, 'N/A') if name_col else 'N/A'
            status = row.get(status_col, 'N/A') if status_col else 'N/A'
            tags = row.get(tag_col, 'N/A') if tag_col else 'N/A'
            epic_url = row.get(url_col, 'N/A') if url_col else 'N/A'
            git_url = row.get(git_col, 'N/A') if git_col else 'N/A'
            
            release_data.append({
                'Release': release,
                'Epic ID': epic_ref,
                'Epic Description': epic_name,
                'Status': status,
                'Epic Link': epic_url,
                'GitHub URL': git_url,
                'Epic Tags': tags
            })
    
    return pd.DataFrame(release_data)

def create_pdf_report(df, status_counts, tag_df, squad_counts, company_df, blocked_df, release_df, report_title, pdf_filename):
    """Create comprehensive PDF report."""
    
    with PdfPages(pdf_filename) as pdf:
        # Title Page
        fig_title = plt.figure(figsize=(11, 8.5))
        ax_title = fig_title.add_subplot(111)
        ax_title.axis('off')
        
        report_date = datetime.now().strftime('%B %d, %Y')
        date_text = f"Generated on: {report_date}"
        
        ax_title.text(0.5, 0.6, report_title, 
                     ha='center', va='center', fontsize=24, fontweight='bold',
                     transform=ax_title.transAxes)
        
        ax_title.plot([0.15, 0.85], [0.57, 0.57], 'k-', linewidth=2, transform=ax_title.transAxes)
        
        ax_title.text(0.5, 0.5, date_text,
                     ha='center', va='center', fontsize=14,
                     transform=ax_title.transAxes)
        
        summary_text = f"Total Epics Analyzed: {len(df)}\n"
        summary_text += f"Epic Status Categories: {len(status_counts)}\n"
        summary_text += f"Squad Tags: {len(squad_counts)}\n"
        summary_text += f"Epics with Company Association: {len(company_df)}"
        if not blocked_df.empty:
            summary_text += f"\n⚠️ Blocked Epics: {len(blocked_df)}"
        
        ax_title.text(0.5, 0.35, summary_text,
                     ha='center', va='center', fontsize=12,
                     transform=ax_title.transAxes,
                     bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.3))
        
        pdf.savefig(fig_title, bbox_inches='tight')
        plt.close()
        
        # Page 1: Epic Status Distribution
        if not status_counts.empty:
            fig = plt.figure(figsize=(14, 10))
            ax = fig.add_subplot(111)
            
            bars = ax.bar(range(len(status_counts)), status_counts.values, 
                         color='steelblue', edgecolor='black', linewidth=1.5)
            
            ax.set_xlabel('Epic Status', fontsize=12, fontweight='bold')
            ax.set_ylabel('Count', fontsize=12, fontweight='bold')
            ax.set_title('Epic Status Distribution', fontsize=14, fontweight='bold', pad=20)
            ax.set_xticks(range(len(status_counts)))
            ax.set_xticklabels(status_counts.index, rotation=45, ha='right')
            
            for bar, value in zip(bars, status_counts.values):
                ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5,
                       str(value), ha='center', va='bottom', fontweight='bold')
            
            ax.grid(axis='y', alpha=0.3, linestyle='--')
            plt.tight_layout()
            pdf.savefig(fig, bbox_inches='tight')
            plt.close()
        
        # Page 2: 3D-style Pie Chart
        if not squad_counts.empty:
            fig = plt.figure(figsize=(12, 10))
            ax1 = fig.add_subplot(111)
            
            colors = plt.cm.Set3(range(len(squad_counts)))
            explode = [0.05 if i == 0 else 0 for i in range(len(squad_counts))]
            
            wedges, texts, autotexts = ax1.pie(squad_counts.values, 
                                               labels=squad_counts.index,
                                               autopct='%1.1f%%',
                                               colors=colors,
                                               startangle=90,
                                               explode=explode,
                                               shadow=True,
                                               textprops={'fontsize': 9})
            
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
            
            # Page 3: Table
            fig = plt.figure(figsize=(11, 8.5))
            ax2 = fig.add_subplot(111)
            ax2.axis('off')
            
            table_data = [['Squad', 'Epic Count']]
            for squad, count in squad_counts.items():
                table_data.append([squad, str(count)])
            
            table = ax2.table(cellText=table_data, cellLoc='left',
                            loc='center', colWidths=[0.7, 0.3])
            table.auto_set_font_size(False)
            table.set_fontsize(10)
            table.scale(1, 2)
            
            for i in range(2):
                table[(0, i)].set_facecolor('#366092')
                table[(0, i)].set_text_props(weight='bold', color='white')
            
            for i in range(1, len(table_data)):
                for j in range(2):
                    if i % 2 == 0:
                        table[(i, j)].set_facecolor('#f0f0f0')
            
            ax2.set_title('Epic Count by Squad', fontsize=16, fontweight='bold', pad=20)
            
            plt.tight_layout()
            pdf.savefig(fig, bbox_inches='tight')
            plt.close()
        
        # Page 4: Company Association Bar Chart
        if not company_df.empty:
            fig = plt.figure(figsize=(14, 10))
            ax = fig.add_subplot(111)
            
            company_counts = company_df['Company Association'].value_counts().sort_values(ascending=False)
            
            bars = ax.bar(range(len(company_counts)), company_counts.values,
                         color='steelblue', edgecolor='black', linewidth=1.5)
            
            ax.set_xlabel('Company Association', fontsize=12, fontweight='bold')
            ax.set_ylabel('Count of Epics', fontsize=12, fontweight='bold')
            ax.set_title(f'Epics by Company Association (Total: {len(company_df)} epics)', 
                        fontsize=14, fontweight='bold', pad=20)
            
            ax.set_xticks(range(len(company_counts)))
            ax.set_xticklabels(company_counts.index, rotation=45, ha='right', fontsize=10)
            
            for bar, value in zip(bars, company_counts.values):
                ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5,
                       str(value), ha='center', va='bottom', fontweight='bold', fontsize=10)
            
            ax.grid(axis='y', alpha=0.3, linestyle='--')
            plt.tight_layout()
            pdf.savefig(fig, bbox_inches='tight')
            plt.close()
            
            # Company Association Details
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
                
                ax.text(0.05, y_position, f"Epic: {epic_ref}",
                       ha='left', va='top', fontsize=10, fontweight='bold',
                       transform=ax.transAxes)
                y_position -= line_height
                
                ax.text(0.08, y_position, f"URL: {epic_url}",
                       ha='left', va='top', fontsize=8, color='blue',
                       transform=ax.transAxes, style='italic')
                y_position -= line_height
                
                ax.text(0.08, y_position, f"Company: {company}",
                       ha='left', va='top', fontsize=8,
                       transform=ax.transAxes)
                y_position -= line_height * 1.5
                
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
        
        # Blockers Page (if any)
        if not blocked_df.empty:
            # Use landscape orientation for better URL display
            fig = plt.figure(figsize=(14, 8.5))
            ax = fig.add_subplot(111)
            ax.axis('off')
            
            # Title with warning symbol
            ax.text(0.5, 0.95, '⚠️ BLOCKERS ⚠️',
                   ha='center', va='top', fontsize=20, fontweight='bold',
                   color='red', transform=ax.transAxes)
            
            ax.text(0.5, 0.90, f'Total Blocked Epics: {len(blocked_df)}',
                   ha='center', va='top', fontsize=14, fontweight='bold',
                   transform=ax.transAxes)
            
            # Create table with full URLs (no truncation)
            table_data = [['Epic Ref', 'Epic Name', 'Epic URL', 'Git URL', 'Epic Tags']]
            
            for idx, row in blocked_df.iterrows():
                epic_ref = str(row.get('Epic Reference', 'N/A'))
                epic_name = str(row.get('Epic Name', 'N/A'))
                epic_url = str(row.get('Epic URL', 'N/A'))
                git_url = str(row.get('Git URL', 'N/A'))
                epic_tags = str(row.get('Epic Tags', 'N/A'))
                
                table_data.append([epic_ref, epic_name, epic_url, git_url, epic_tags])
            
            # Adjust column widths to accommodate full URLs
            table = ax.table(cellText=table_data, cellLoc='left',
                           loc='center', colWidths=[0.10, 0.20, 0.28, 0.28, 0.14])
            table.auto_set_font_size(False)
            table.set_fontsize(7)  # Smaller font to fit URLs
            table.scale(1, 1.8)
            
            # Style header row
            for i in range(5):
                table[(0, i)].set_facecolor('#FF6B6B')
                table[(0, i)].set_text_props(weight='bold', color='white', fontsize=8)
            
            # Alternate row colors and enable text wrapping
            for i in range(1, len(table_data)):
                for j in range(5):
                    cell = table[(i, j)]
                    if i % 2 == 0:
                        cell.set_facecolor('#FFE5E5')
                    # Enable text wrapping for URLs
                    cell.set_text_props(wrap=True)
            
            plt.tight_layout()
            pdf.savefig(fig, bbox_inches='tight')
            plt.close()
        
        # Release Analysis Pages (at the end)
        if not release_df.empty:
            # Define status colors
            status_colors = {
                'New': '#87CEEB',  # Sky blue
                'In Development': '#FFD700',  # Gold
                'In Design': '#DDA0DD',  # Plum
                'Dev Complete': '#90EE90',  # Light green
                'Shipped': '#32CD32',  # Lime green
                'Ready for Development': '#FFA500',  # Orange
                'Ready for Design': '#FF69B4',  # Hot pink
                'Under Consideration': '#D3D3D3',  # Light gray
                'Blocked': '#FF0000',  # Red
                'UX Design Delivered': '#9370DB'  # Medium purple
            }
            
            # Group by release
            for release in release_df['Release'].unique():
                release_epics = release_df[release_df['Release'] == release]
                
                # Create landscape page for better table display
                fig = plt.figure(figsize=(14, 8.5))
                ax = fig.add_subplot(111)
                ax.axis('off')
                
                ax.text(0.5, 0.95, f'Release: {release}',
                       ha='center', va='top', fontsize=16, fontweight='bold',
                       transform=ax.transAxes)
                
                ax.text(0.5, 0.90, f'Total Epics: {len(release_epics)}',
                       ha='center', va='top', fontsize=12,
                       transform=ax.transAxes)
                
                # Create table with new columns
                table_data = [['Epic ID', 'Epic Description', 'Status', 'Epic Link', 'GitHub URL', 'Epic Tags']]
                
                for idx, row in release_epics.iterrows():
                    epic_id = str(row.get('Epic ID', 'N/A'))
                    epic_desc = str(row.get('Epic Description', 'N/A'))
                    status = str(row.get('Status', 'N/A'))
                    epic_link = str(row.get('Epic Link', 'N/A'))
                    git_url = str(row.get('GitHub URL', 'N/A'))
                    tags = str(row.get('Epic Tags', 'N/A'))
                    
                    # Truncate long fields
                    if len(epic_desc) > 40:
                        epic_desc = epic_desc[:37] + '...'
                    if len(tags) > 30:
                        tags = tags[:27] + '...'
                    
                    table_data.append([epic_id, epic_desc, status, epic_link, git_url, tags])
                
                # Split into multiple pages if needed (max 20 rows per page for more columns)
                rows_per_page = 20
                num_pages = (len(table_data) - 1 + rows_per_page - 1) // rows_per_page
                
                for page_num in range(num_pages):
                    if page_num > 0:
                        # Create new page
                        pdf.savefig(fig, bbox_inches='tight')
                        plt.close()
                        
                        fig = plt.figure(figsize=(14, 8.5))
                        ax = fig.add_subplot(111)
                        ax.axis('off')
                        
                        ax.text(0.5, 0.95, f'Release: {release} (continued {page_num + 1})',
                               ha='center', va='top', fontsize=16, fontweight='bold',
                               transform=ax.transAxes)
                    
                    start_idx = page_num * rows_per_page + 1
                    end_idx = min((page_num + 1) * rows_per_page + 1, len(table_data))
                    
                    page_data = [table_data[0]] + table_data[start_idx:end_idx]
                    
                    # Create table with 6 columns
                    table = ax.table(cellText=page_data, cellLoc='left',
                                   loc='center', colWidths=[0.08, 0.20, 0.12, 0.22, 0.22, 0.16])
                    table.auto_set_font_size(False)
                    table.set_fontsize(7)
                    table.scale(1, 1.8)
                    
                    # Style header row
                    for i in range(6):
                        table[(0, i)].set_facecolor('#366092')
                        table[(0, i)].set_text_props(weight='bold', color='white', fontsize=7)
                    
                    # Color code status cells and enable text wrapping for URLs
                    for i in range(1, len(page_data)):
                        status = page_data[i][2]  # Status is now column 2
                        color = status_colors.get(status, '#FFFFFF')
                        table[(i, 2)].set_facecolor(color)
                        
                        # Enable text wrapping for URL columns
                        table[(i, 3)].set_text_props(wrap=True)  # Epic Link
                        table[(i, 4)].set_text_props(wrap=True)  # GitHub URL
                        
                        # Alternate row colors for other columns
                        if i % 2 == 0:
                            for j in [0, 1, 5]:  # Epic ID, Description, Tags
                                table[(i, j)].set_facecolor('#f0f0f0')
                
                plt.tight_layout()
                pdf.savefig(fig, bbox_inches='tight')
                plt.close()
    
    print(f"✓ PDF report saved: {pdf_filename}")
    return pdf_filename

def analyze_report(filename, report_title, output_prefix):
    """Analyze a single report."""
    print(f"\n{'='*70}")
    print(f"Analyzing: {report_title}")
    print(f"{'='*70}")
    
    if not os.path.exists(filename):
        print(f"Error: File not found: {filename}")
        return None
    
    print(f"Reading Excel file: {filename}")
    df = pd.read_excel(filename)
    print(f"Total records: {len(df)}")
    
    status_counts, status_col = analyze_epic_status(df)
    tag_df, squad_counts = analyze_epic_tags(df)
    company_df = analyze_company_association(df)
    blocked_df = get_blocked_epics(df, status_col)
    release_df = analyze_by_release(df)
    
    if not status_counts.empty:
        print("\nEpic Status Distribution:")
        for status, count in status_counts.items():
            print(f"  {status}: {count}")
    
    if not squad_counts.empty:
        print(f"\nTop 10 Squad Tags:")
        for squad, count in squad_counts.head(10).items():
            print(f"  {squad}: {count}")
    
    print(f"\nEpics with Company Association: {len(company_df)}")
    
    if not blocked_df.empty:
        print(f"⚠️  Blocked Epics: {len(blocked_df)}")
    
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    pdf_filename = f'{output_prefix}_epic_analysis_{timestamp}.pdf'
    
    create_pdf_report(df, status_counts, tag_df, squad_counts, company_df, blocked_df, release_df, report_title, pdf_filename)
    
    return pdf_filename

def main():
    """Main function - runs both analyses."""
    print("="*70)
    print("DI AHA ANALYSER - Master Script")
    print("="*70)
    
    reports = [
        {
            'filename': 'aha_list_features_260325061137.xlsx',
            'title': 'IKC 2026 Epic Planning & Execution Report',
            'prefix': 'ikc'
        },
        {
            'filename': 'aha_list_features_260326054547.xlsx',
            'title': 'Lineage 2026 Epic Planning & Execution Report',
            'prefix': 'lineage'
        }
    ]
    
    generated_reports = []
    
    for report in reports:
        try:
            pdf_file = analyze_report(
                report['filename'],
                report['title'],
                report['prefix']
            )
            if pdf_file:
                generated_reports.append(pdf_file)
        except Exception as e:
            print(f"\nError processing {report['title']}: {e}")
            import traceback
            traceback.print_exc()
    
    print("\n" + "="*70)
    print("Analysis Complete!")
    print("="*70)
    print(f"\nGenerated {len(generated_reports)} reports:")
    for report in generated_reports:
        print(f"  ✓ {report}")
    print("="*70)

if __name__ == "__main__":
    main()

# Made with Bob
