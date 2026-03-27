# Executive AHA Epic Analysis Reports

## Overview

Comprehensive executive reporting system that generates professional PDF and PowerPoint presentations analyzing epic planning and execution for IKC and Lineage teams. Features include status tracking, squad analysis, company associations, blocker identification, and detailed release-level analysis with clickable links.

## Quick Start

```bash
python3 executive_updates_aha_analysis.py
```

This single command generates all reports in ~30-35 seconds:
- 2 detailed PDF reports (IKC and Lineage)
- 1 combined PowerPoint presentation

## Generated Files

### PDF Reports
- `ikc_epic_analysis_YYYYMMDD_HHMMSS.pdf` (~155KB)
- `lineage_epic_analysis_YYYYMMDD_HHMMSS.pdf` (~121KB)

### PowerPoint Presentation
- `aha_epic_analysis_combined_YYYYMMDD_HHMMSS.pptx` (~141KB, 42+ slides)

## Prerequisites

```bash
pip install pandas matplotlib openpyxl python-pptx python-dotenv
```

## Input Files

Two Excel files exported from Aha!:
1. `aha_list_features_260325061137.xlsx` - IKC epic data (213 epics)
2. `aha_list_features_260326054547.xlsx` - Lineage epic data (187 epics)

### Required Excel Columns
- Epic URL
- Epic reference #
- Epic name
- Epic type
- Company Association
- Epic status
- Epic product value score
- Release name
- Epic tags
- Github enterprise html_url
- Epic initial estimate
- Epic created date
- Product Management owner
- Development owner
- Epic assigned to

## Report Contents

### PDF Reports Include

#### 1. Title Page
- Report name (underlined)
- Generation date
- Summary statistics (total epics, status categories, squad tags, company associations, blocked epics)

#### 2. Epic Status Distribution
- Bar chart showing status breakdown
- Counts for each status category

#### 3. Squad Distribution
- 3D-style pie chart with shadow and explode effects
- Detailed table with all squad tags and counts
- Automatic pagination for large datasets

#### 4. Company Association Analysis
- Bar chart showing distribution by company
- Detailed listings with Epic Reference, Epic URL, and Company name
- Multiple pages if needed (15 rows per page)

#### 5. Blocked Epics (if applicable)
- Red-themed warning page
- Table with: Epic Ref, Epic Name, Epic URL, Git URL, Epic Tags
- Full URLs displayed (no truncation)
- Landscape orientation for better readability

#### 6. Release Analysis (NEW)
- **Grouped by Release Name**
- **6-column table**:
  - Epic ID
  - Epic Description
  - Status (color-coded)
  - Epic Link (full URL)
  - GitHub URL (full URL)
  - Epic Tags
- **Color-coded status cells** (10 distinct colors)
- **20 rows per page**
- **Landscape orientation**
- **Text wrapping for URLs**

### PowerPoint Presentation Includes

#### Section 1: Executive Summary
- Title slide with report date
- Key metrics for both IKC and Lineage teams

#### Section 2: IKC Analysis
- Status distribution chart
- Complete squad tags table (all squads, paginated)
- Company association donut chart with companion table
- Company association detailed listings
- Blocked epics table (if applicable)
- **Release analysis with clickable URLs**

#### Section 3: Lineage Analysis
- Status distribution chart
- Complete squad tags table (all squads, paginated)
- Company association donut chart with companion table
- Company association detailed listings
- Blocked epics table (if applicable)
- **Release analysis with clickable URLs**

#### Section 4: Comparison
- Side-by-side metrics comparison chart

## Key Features

### Release Analysis (Enhanced)

**Table Columns**:
1. **Epic ID** - Reference number (0.8" width)
2. **Epic Description** - Full epic name (1.8" width)
3. **Status** - Color-coded cell (1.2" width)
4. **Epic Link** - Full Aha! URL, clickable (2.2" width)
5. **GitHub URL** - Full GitHub enterprise URL, clickable (2.2" width)
6. **Epic Tags** - Squad assignments (1.2" width)

**PowerPoint Features**:
- ✅ **Clickable URLs** - Blue, underlined hyperlinks
- ✅ **Full URLs** - No truncation
- ✅ **Word wrap** - URLs wrap within cells
- ✅ **15 rows per slide** - Optimized for readability
- ✅ **Professional formatting** - Alternating row colors

**PDF Features**:
- ✅ **Full URLs** - Complete URLs displayed
- ✅ **Text wrapping** - URLs wrap within cells
- ✅ **20 rows per page** - Landscape orientation
- ✅ **Color-coded status** - Visual status tracking

### Status Color Coding

| Status | Color | RGB |
|--------|-------|-----|
| New | Sky Blue | (135, 206, 235) |
| In Development | Gold | (255, 215, 0) |
| In Design | Plum | (221, 160, 221) |
| Dev Complete | Light Green | (144, 238, 144) |
| Shipped | Lime Green | (50, 205, 50) |
| Ready for Development | Orange | (255, 165, 0) |
| Ready for Design | Hot Pink | (255, 105, 180) |
| Under Consideration | Light Gray | (211, 211, 211) |
| Blocked | Red | (255, 0, 0) |
| UX Design Delivered | Medium Purple | (147, 112, 219) |

### Company Association Display

**PowerPoint Layout**:
- **Left side**: Simple donut chart (4.5" × 5")
- **Right side**: Two-column table (Company | Count)
- **Legend**: Bottom of donut chart
- **Table styling**: Blue header, alternating gray rows

### Automatic Features

- ✅ **Timestamp in filenames** - Easy version tracking
- ✅ **Blocker detection** - Automatic identification and highlighting
- ✅ **Multi-page support** - Automatic pagination for large datasets
- ✅ **Error handling** - Comprehensive error reporting
- ✅ **Execution summary** - Timing and status information

## Usage Examples

### Weekly Executive Updates
```bash
python3 executive_updates_aha_analysis.py
```

### Monthly Planning Reviews
```bash
python3 executive_updates_aha_analysis.py
```

### Ad-hoc Analysis
```bash
python3 executive_updates_aha_analysis.py
```

## Script Architecture

### Master Script
[`executive_updates_aha_analysis.py`](executive_updates_aha_analysis.py)
- Orchestrates both PDF and PowerPoint generation
- Provides execution summary with timing
- Handles errors gracefully

### PDF Generator
[`di_aha_analyser.py`](di_aha_analyser.py)
- Generates detailed PDF reports
- Includes all analysis sections
- Creates color-coded tables
- Handles release analysis with full URLs

### PowerPoint Generator
[`create_aha_ppt.py`](create_aha_ppt.py)
- Creates executive presentation
- Generates donut charts and tables
- Makes URLs clickable
- Handles release analysis with hyperlinks

## Customization

### Modifying Input Files

Edit file paths in underlying scripts:

**di_aha_analyser.py** (lines 449-460):
```python
reports = [
    {
        'filename': 'your_ikc_file.xlsx',
        'title': 'IKC 2026 Epic Planning & Execution Report',
        'prefix': 'ikc'
    },
    {
        'filename': 'your_lineage_file.xlsx',
        'title': 'Lineage 2026 Epic Planning & Execution Report',
        'prefix': 'lineage'
    }
]
```

**create_aha_ppt.py** (lines 223-226):
```python
ikc_data = analyze_data('your_ikc_file.xlsx')
lineage_data = analyze_data('your_lineage_file.xlsx')
```

### Adjusting Rows Per Page/Slide

**PDF** (di_aha_analyser.py, line 497):
```python
rows_per_page = 20  # Change to desired number
```

**PowerPoint** (create_aha_ppt.py, line 565):
```python
rows_per_slide = 15  # Change to desired number
```

### Changing Status Colors

Edit the `status_colors` dictionary in both scripts:
```python
status_colors = {
    'New': RGBColor(135, 206, 235),  # Your custom color
    # ... add more
}
```

## Troubleshooting

### Missing Input Files
```
Error: File not found: aha_list_features_260325061137.xlsx
```
**Solution**: Ensure Excel files are in the current directory

### Import Errors
```
ModuleNotFoundError: No module named 'pandas'
```
**Solution**: Install required packages:
```bash
pip install pandas matplotlib openpyxl python-pptx
```

### Script Execution Fails
Check individual scripts work:
```bash
python3 di_aha_analyser.py
python3 create_aha_ppt.py
```

### URLs Not Clickable
- Ensure URLs are valid (start with http:// or https://)
- Check that cell data is not 'N/A'
- Verify PowerPoint version supports hyperlinks

## Performance

- **Typical execution time**: 30-35 seconds
- **IKC dataset**: ~213 epics
- **Lineage dataset**: ~187 epics
- **Total output size**: ~417KB (2 PDFs + 1 PPTX)
- **Total slides**: 42+ (varies with number of releases)

## Output Statistics

### IKC Report
- Total Epics: 213
- Status Categories: 9
- Squad Tags: ~40
- Company Associations: 81
- Blocked Epics: 0
- Releases: Multiple (varies)

### Lineage Report
- Total Epics: 187
- Status Categories: 10
- Squad Tags: ~30
- Company Associations: 29
- Blocked Epics: 2
- Releases: Multiple (varies)

## Automation

### Scheduled Execution

**Cron (Linux/Mac)**:
```bash
# Run every Monday at 9 AM
0 9 * * 1 cd /path/to/project && python3 executive_updates_aha_analysis.py
```

**Task Scheduler (Windows)**:
Create a scheduled task to run the script weekly

### CI/CD Integration

```yaml
- name: Generate Executive Reports
  run: python3 executive_updates_aha_analysis.py
  
- name: Upload Reports
  uses: actions/upload-artifact@v2
  with:
    name: aha-reports
    path: |
      *_epic_analysis_*.pdf
      *_epic_analysis_*.pptx
```

## Version History

- **v2.0** (2026-03-26): Enhanced release analysis
  - Added Epic Description, Epic Link, GitHub URL columns
  - Made URLs clickable in PowerPoint
  - Full URL display (no truncation)
  - Improved table formatting
  
- **v1.5** (2026-03-26): Release analysis added
  - Color-coded status tracking
  - Grouped by release name
  - Automatic pagination
  
- **v1.0** (2026-03-26): Initial release
  - Combined PDF and PowerPoint generation
  - Status, squad, and company analysis
  - Blocker detection

## Related Documentation

- [`README_AHA_EPIC_ANALYSIS.md`](README_AHA_EPIC_ANALYSIS.md) - PDF report details
- [`README_AHA_PPT_GENERATOR.md`](README_AHA_PPT_GENERATOR.md) - PowerPoint details
- [`README_EXECUTIVE_UPDATES.md`](README_EXECUTIVE_UPDATES.md) - Master script details

## Support

For issues or questions:
1. Check input files are current and properly formatted
2. Verify all required columns exist in Excel files
3. Review error messages in execution output
4. Ensure all dependencies are installed
5. Contact development team for assistance

## Best Practices

1. **Regular Updates**: Run weekly before leadership meetings
2. **Version Control**: Keep generated reports in dated folders
3. **Data Validation**: Verify Excel exports are complete
4. **Backup**: Maintain copies of input files
5. **Review**: Always review generated reports before distribution

---

**Generated by**: Bob AI Assistant  
**Last Updated**: March 26, 2026  
**Version**: 2.0  
**Scripts**: 
- [`executive_updates_aha_analysis.py`](executive_updates_aha_analysis.py)
- [`di_aha_analyser.py`](di_aha_analyser.py)
- [`create_aha_ppt.py`](create_aha_ppt.py)