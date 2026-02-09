# CUP Metadata Processor

A unified web-based application for converting Excel metadata to XML and validating Cambridge University Press (CUP) book specification XML files.

## Features

### Excel → XML Converter
- **Batch Processing**: Convert multiple Excel rows into individual XML files
- **Automatic Column Mapping**: Maps Excel column headers to standardized XML attribute names
- **Value Remapping**: Automatically transforms specific values (paper types, lamination)
- **Spine Calculation**: Automatically calculates spine size based on extent, paper type, and binding style
- **GSM Extraction**: Extracts paper weight from paper names
- **ZIP Download**: Packages all XML files into a single ZIP archive
- **ISBN-based Naming**: XML files are named using ISBN values
- **Summary Statistics**: View record counts by binding, colour, and quality

### XML Validator
- **Multi-file Upload**: Drag & drop or browse multiple XML files
- **Real-time Validation**: Instant validation results with visual indicators
- **8 Comprehensive Checks**: Covers all specification requirements
- **Summary Reports**: Download validation results as text reports
- **Failed-Only Reports**: Download summary of only failed validations
- **Copy to Clipboard**: Quick copy of individual validation results
- **Filter Results**: Toggle to show only failed validations
- **Compact Interface**: Clean, easy-to-read results display

### Integrated Workflow
- **Seamless Processing**: Convert Excel → Generate XMLs → Validate → Download
- **One-Click Validation**: "Validate Generated XMLs" button automatically validates converted files
- **Tab-based Interface**: Clean separation of conversion and validation workflows

## Installation

1. Download all files to a single directory:
   - `cup-metadata-processor.html`
   - `cup-processor-config.js`
   - `cup-processor-converter.js`
   - `cup-processor-validator.js`
   - `cup-processor-app.js`
   - `cup-processor-style.css`

2. Open `cup-metadata-processor.html` in a modern web browser

No server or build process required - runs entirely in the browser.

## Usage

### Tab 1: Excel → XML Converter

#### 1. Prepare Your Excel File

Your Excel file should have the following columns:

| Original Column Name | XML Attribute | Description |
|---------------------|---------------|-------------|
| Master ISBN 13 | isbn | 13-digit ISBN number |
| full_title_for_export | title | Book title |
| pod_xml_TrimHeight | trim_height | Trim height in mm |
| pod_xml_TrimWidth | trim_width | Trim width in mm |
| pod_ExtentMod2 | extent | Number of pages |
| pod_xml_TJ_PaperStock | paper | Paper type |
| pod_Colourspace | colour | Colour type (Mono/Colour) |
| pod_xml_TJ_QualityOption | quality | Quality (Standard/Premium) |
| pod_ColourBook_fullBleed | bleed | Bleed option (Y/N) |
| pod_TJ_BindType_c | binding_style | Binding type (Cased/Limp) |
| pod_CoverLaminate | lamination | Lamination type |
| pod_xml_Jacket | jacket | Jacket option (Y/N) |
| pod_xml_TJ_JacketLamination | jacket_lamination | Jacket lamination type |
| pod_xml_TJ_FlapSize | flap_size | Flap size |
| pod_xml_TJ_WibalinColour | wibalin_colour | Wibalin colour |
| pod_xml_SetISBN | set_isbn | Set ISBN |

#### 2. Upload and Process

1. Click "Choose Excel File" and select your .xlsx or .xls file
2. Review the summary statistics showing record counts
3. Click "Download All XMLs as ZIP" to download generated files
4. **Optional**: Click "Validate Generated XMLs" to automatically validate in Tab 2

#### 3. Data Transformations

**Paper Type Remapping:**
- `Cream 80gsm` → `CUP MunkenPure 80 gsm`
- `White 80gsm` → `Navigator 80 gsm`
- `Matte Coated 90gsm` → `Clairjet 90 gsm` (Standard) or `Magno Matt 90 gsm` (Premium)

**Lamination Remapping:**
- `Matte` → `Matt`

**Spine Calculation:**
```
Base Spine = (Extent × Grammage × Volume) / 20000
If Cased binding: Add 4mm
Final Spine = Base Spine (+ 4 if Cased)
```

**Paper Specifications:**

| Paper Type | Grammage | Volume |
|-----------|----------|--------|
| CUP MunkenPure 80 gsm | 80 | 13 |
| Navigator 80 gsm | 80 | 12.5 |
| Clairjet 90 gsm | 90 | 10 |
| Magno Matt 90 gsm | 90 | 10 |

#### 4. XML Output Structure

```xml
<?xml version="1.0" encoding="UTF-8"?>
<record>
  <isbn>9781234567890</isbn>
  <title>Book Title</title>
  <trim_height>229</trim_height>
  <trim_width>152</trim_width>
  <extent>662</extent>
  <spine>50</spine>
  <paper>CUP MunkenPure 80 gsm</paper>
  <gsm>80</gsm>
  <colour>Mono</colour>
  <quality>Standard</quality>
  <bleed>N</bleed>
  <binding_style>Cased</binding_style>
  <lamination>Matt</lamination>
  <jacket>N</jacket>
  <jacket_lamination></jacket_lamination>
  <flap_size></flap_size>
  <wibalin_colour></wibalin_colour>
  <set_isbn></set_isbn>
</record>
```

### Tab 2: XML Validator

#### 1. Upload XML Files

- Drag and drop XML files onto the upload area, or
- Click "Browse Files" to select files from your computer
- Multiple files can be processed at once

#### 2. View Results

- Each file displays as a card showing pass/fail status
- **Green border** = all validations passed
- **Red border** = one or more validations failed
- Detailed validation messages appear in each card

#### 3. Filter and Export

- Use "Show failed only" toggle to display only files with errors
- Click "Download Full Summary" to export all results as a text file
- Click "Download Failed Summary" to export only failed validations
- Click clipboard icon on individual cards to copy results

#### 4. Clear and Restart

- Click "Clear Results" to remove all files and start over

## Validation Rules

### Required Fields
- ISBN, Trim Height, Trim Width, Extent, Paper, Colour, Quality, Binding Style

### Valid Values

**Binding Style:**
- Cased
- Limp

**Paper Types:**
- CUP MunkenPure 80 gsm
- Navigator 80 gsm
- Clairjet 90 gsm
- Magno Matt 90 gsm

**Colour:**
- Mono
- Colour

**Quality/Route:**
- Standard
- Premium

**Trim Sizes (width x height in mm):**
- 140x216
- 152x229
- 156x234
- 170x244
- 189x246
- 178x254
- 203x254
- 216x280

### Paper Compatibility Rules

| Paper Type | Allowed Colour | Allowed Route |
|-----------|---------------|---------------|
| CUP MunkenPure 80 gsm | Mono only | Standard only |
| Navigator 80 gsm | Mono only | Standard only |
| Clairjet 90 gsm | Colour only | Standard only |
| Magno Matt 90 gsm | Mono or Colour | Premium only |

## Workflow Example

1. **Convert Excel to XML**
   - Upload Excel file in Tab 1
   - Review summary statistics
   - Click "Validate Generated XMLs"

2. **Automatic Validation**
   - Application switches to Tab 2
   - All generated XMLs are automatically validated
   - Results displayed immediately

3. **Review and Export**
   - Filter to show only failed validations if needed
   - Download summary report
   - Fix any issues in source Excel
   - Repeat process

## Browser Compatibility

- Chrome 90+ (recommended)
- Firefox 88+
- Safari 14+
- Edge 90+

## Technical Details

### Libraries Used
- **SheetJS (xlsx.js)**: Excel file parsing
- **JSZip**: ZIP file creation
- **Bootstrap 5**: UI framework
- **Bootstrap Icons**: Icons

### Security
- All processing happens locally in your browser
- No data is uploaded to any server
- No external API calls are made

### File Structure
- `cup-metadata-processor.html` - Main HTML interface with tab navigation
- `cup-processor-config.js` - Shared configuration and validation rules
- `cup-processor-converter.js` - Excel to XML conversion logic
- `cup-processor-validator.js` - XML validation logic and CUPValidator class
- `cup-processor-app.js` - Application initialization
- `cup-processor-style.css` - Custom styling for both modules

## Troubleshooting

### Converter Issues

**File Won't Upload**
- Ensure file is .xlsx or .xls format
- Check file isn't corrupted
- Try with a smaller file first

**Missing Spine Values**
- Ensure paper type exactly matches supported types
- Verify extent and binding_style columns have valid data
- Check paper names are being properly mapped

**Incorrect Values in XML**
- Verify Excel column headers match expected names exactly
- Check for leading/trailing spaces in data
- Ensure data types are correct (numbers for extent, trim_height, etc.)

### Validator Issues

**Files Won't Validate**
- Ensure files are valid XML format
- Check for proper `<record>` root element
- Verify all required fields are present

**Validation Always Fails**
- Check that values exactly match expected formats
- Verify paper/colour/route combinations are compatible
- Ensure trim sizes are from the approved list

### General Issues

**ZIP Download Fails**
- Try with fewer records
- Check browser console (F12) for errors
- Ensure JavaScript is enabled

**Application Won't Load**
- Verify all files are in the same directory
- Check browser console for missing file errors
- Try a different browser

## Support

For issues or questions:
1. Check browser console (F12) for error messages
2. Verify file structure matches requirements
3. Test with sample data first

## Version History

### Version 1.0
- Initial release
- Unified Excel-to-XML converter and XML validator
- Tab-based interface
- Integrated workflow with automatic validation
- Shared configuration between modules
- Complete documentation

---

**Version:** 1.0  
**Last Updated:** February 2026  
**Created by:** Colin for Cambridge University Press
