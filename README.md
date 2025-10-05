# Excel to XML Batch Converter

A web-based tool for converting Excel spreadsheets into individual XML files with automatic column mapping, value transformation, and spine size calculation.

## Features

- **Batch Processing**: Convert multiple rows from Excel into individual XML files
- **Automatic Column Mapping**: Maps Excel column headers to standardized XML attribute names
- **Value Remapping**: Automatically transforms specific values (e.g., paper types, lamination)
- **Spine Calculation**: Automatically calculates spine size based on extent, paper type, and binding style
- **Statistical Summary**: Displays detailed statistics about your data
- **ZIP Download**: Packages all XML files into a single ZIP archive
- **ISBN-based Naming**: XML files are named using ISBN values

## Getting Started

### Prerequisites

- Modern web browser (Chrome, Firefox, Safari, or Edge)
- No server or installation required - runs entirely in the browser

### Installation

1. Download both files:
   - `index.html`
   - `script.js`

2. Place both files in the same directory

3. Open `index.html` in your web browser

## Usage

### 1. Prepare Your Excel File

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

### 2. Upload Excel File

1. Click the file upload area or drag and drop your Excel file
2. Supported formats: `.xlsx`, `.xls`

### 3. Review Summary

After upload, you'll see a processing summary showing:

- **Total Records**: Number of rows processed
- **Binding Style**: Count of Cased and Limp bindings
- **Quality**: Count of Standard and Premium quality
- **Colour**: Count of Mono and Colour jobs
- **Jobs with Bleeds**: Count of jobs with bleed
- **Jackets**: Count of jobs with jackets
- **Lamination Types**: Breakdown of all lamination types

### 4. Download XML Files

Click the "Download All XMLs as ZIP" button to download a ZIP archive containing all XML files.

## Data Transformations

### Column Name Mapping

All column headers are automatically mapped from the Excel format to clean XML attribute names as shown in the table above.

### Value Remapping

The following values are automatically transformed:

#### Paper Types:
- `Cream 80gsm` → `Munken 80 gsm`
- `White 80gsm` → `Navigator 80 gsm`
- `Matte Coated` → `LetsGo Silk 90 gsm`

#### Lamination:
- `Matte` → `Matt`

### Spine Calculation

The spine size is automatically calculated using the formula:

```
Base Spine = (Extent × Grammage × Volume) / 20000
If Cased binding: Add 4mm
Final Spine = Base Spine (+ 4 if Cased)
```

**Paper Specifications:**

| Paper Type | Grammage | Volume |
|-----------|----------|--------|
| Munken 80 gsm | 80 | 17.5 |
| Navigator 80 gsm | 80 | 12.5 |
| LetsGo Silk 90 gsm | 90 | 10 |

**Example:**
- 662 pages, Munken 80 gsm, Cased binding
- Base = (662 × 80 × 17.5) / 20000 = 46mm
- Final = 46 + 4 = **50mm**

## XML Output Structure

Each row generates an XML file with the following structure:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<record>
  <isbn>9781234567890</isbn>
  <title>Book Title</title>
  <trim_height>229</trim_height>
  <trim_width>152</trim_width>
  <extent>662</extent>
  <spine>50</spine>
  <paper>Munken 80 gsm</paper>
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

## File Naming

XML files are automatically named using the ISBN value:

- Format: `[ISBN].xml`
- Example: `9781234567890.xml`

If an ISBN is missing or invalid, files use a numbered format: `record_0001.xml`, `record_0002.xml`, etc.

## Column Order

The XML attributes appear in the following order:

1. isbn
2. title
3. trim_height
4. trim_width
5. extent
6. **spine** (calculated)
7. paper
8. colour
9. quality
10. bleed
11. binding_style
12. lamination
13. jacket
14. jacket_lamination
15. flap_size
16. wibalin_colour
17. set_isbn

## Technical Details

### Libraries Used

- **SheetJS (xlsx.js)**: Excel file parsing
- **JSZip**: ZIP file creation
- **Bootstrap 5**: UI framework
- **Bootstrap Icons**: Icons

### Browser Compatibility

- Chrome 90+
- Firefox 88+
- Safari 14+
- Edge 90+

### Security

- All processing happens locally in your browser
- No data is uploaded to any server
- No external API calls are made

## Troubleshooting

### File Won't Upload
- Ensure the file is in `.xlsx` or `.xls` format
- Check that the file isn't corrupted
- Try with a smaller file first

### Missing Spine Values
- Ensure the paper type exactly matches one of the supported types
- Verify extent and binding_style columns have valid data
- Check that paper names are being properly mapped

### Incorrect Values in XML
- Verify your Excel column headers match the expected names exactly
- Check for leading/trailing spaces in data values
- Ensure data types are correct (numbers for extent, trim_height, etc.)

### ZIP Download Fails
- Try with fewer records
- Check browser console for errors
- Ensure JavaScript is enabled

## Support

For issues or questions:
1. Check the browser console (F12) for error messages
2. Verify your Excel file structure matches the requirements
3. Test with a small sample file first

## License

This tool is provided as-is for internal use.

## Version History

### Version 1.0
- Initial release
- Excel to XML conversion
- Automatic column mapping
- Value remapping for paper and lamination
- Spine size calculation
- Statistical summary
- ZIP download functionality