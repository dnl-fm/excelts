# Excel Format Notes

Historical implementation notes about Excel XLSX XML structure. Reference material for understanding the format internals.

---

## Tables

### Sheet
- Add `<tableParts>`

### sheet.rels
- Add ref to `../../tables/table1.xml`

### table1.xml
- Define the table

### Content_Types
- Add table1

### Notes
- In `table1.xml`, the `tableColumns` can define a `dataDxfId` (rel to styles.xml) that specifies default format of new data
- In `table1.xml`, the `tableColumn` name property must match the value in the cell where the column header goes

### Table Properties
- `ref`: Top Left of table
- `name`: table.name
- `displayName`: table.displayName
- `range`: range of data cells (not including headers/footers)
  - `table.ref` includes headers/footers
  - `autoFilter.ref` includes headers only
- `headerRow`: table.headerRowCount
  - default: true, show/hide the headers
- `totalsRow`: `table.totalsRowShown="0"` or `table.totalsRowCount="1"`
  - default: false, show/hide totals rows
- `columns`: tableColumns[n]
  - `id`: sequential number
  - `name`: must correspond to top cell's value
  - `style` -> dataDxfId (inserts inline style into dxfs in styles.xml)
  - `filterButton`: `autoFilter[n].hiddenButton="1"` (shows or hides the filter, default: true)
  - `totalsRowLabel`: for column1
  - `totalsRowFunction`: name of function, one of [none, count, etc]
  - `totalsRowFormula`: if function is 'custom', a formula
  - `totalsRowResult`: the value to show
- `style`:
  - `name`: tableStyleInfo.name
  - `showFirstColumn`: tableStyleInfo.showFirstColumn
  - `showLastColumn`: tableStyleInfo.showLastColumn
  - `showRowStripes`: tableStyleInfo.showRowStripes
  - `showColumnStripes`: tableStyleInfo.showColumnStripes

---

## Page Setup

### Themes
Not supported

### Margins (in inches)
```xml
<pageMargins left="0.25" right="0.25" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
```

### Page Setup Examples
```xml
<pageSetup paperSize="9" orientation="portrait" horizontalDpi="4294967295" verticalDpi="4294967295" r:id="rId1"/>
<pageSetup paperSize="119" pageOrder="overThenDown" orientation="portrait" blackAndWhite="1"
    draft="1" cellComments="atEnd" errors="dash" horizontalDpi="4294967295" verticalDpi="4294967295" r:id="rId1"/>
<pageSetup paperSize="119" scale="95" fitToWidth="2" fitToHeight="3" orientation="landscape" r:id="rId1"/>
```

### Page Setup Attributes
| Attribute | Values |
|-----------|--------|
| orientation | portrait, landscape |
| pageOrder | overThenDown, downThenOver(*) |
| blackAndWhite | true, false(*) |
| draft | true, false(*) |
| cellComments | atEnd, asDisplayed, None(*) |
| errors | dash, blank, NA, displayed(*) |
| scale | number |
| fitToWidth | num pages wide |
| fitToHeight | num pages high |

### Fit to Page Detection
```xml
<sheetPr>
    <pageSetUpPr fitToPage="1"/>
</sheetPr>
```

### Paper Sizes
| Name | Code |
|------|------|
| Letter (*) | - |
| Legal | 5 |
| Executive | 7 |
| A4 | 9 |
| A5 | 11 |
| B5 (JIS) | 13 |
| Envelope #10 | 20 |
| Envelope DL | 27 |
| Envelope C5 | 28 |
| Envelope B5 | 34 |
| Envelope Monarch | 37 |
| Double Japan Postcard Rotated | 82 |
| 16K 197x273 mm | 119 |

---

## Print Options

```xml
<printOptions headings="1" gridLines="1"/>
<printOptions horizontalCentered="1" verticalCentered="1"/>
```

| Option | Values |
|--------|--------|
| headings | true, false(*) |
| gridLines | true, false(*) |
| horizontalCentered | true, false(*) |
| verticalCentered | true, false(*) |

---

## Print Area

### In workbook.xml
```xml
<definedNames>
    <definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$C$11</definedName>
</definedNames>
```

### In app.xml
```xml
<TitlesOfParts>
    <vt:vector size="2" baseType="lpstr">
        <vt:lpstr>Sheet1</vt:lpstr>
        <vt:lpstr>Sheet1!Print_Area</vt:lpstr>
    </vt:vector>
</TitlesOfParts>
```

---

## Manual Page Breaks

### In sheet.xml
```xml
<rowBreaks count="2" manualBreakCount="2">
    <brk id="5" max="16383" man="1"/>  <!-- underneath row 5 -->
    <brk id="7" max="16383" man="1"/>
</rowBreaks>
<colBreaks count="3" manualBreakCount="3">
    <brk id="3" max="1048575" man="1"/>
    <brk id="6" max="1048575" man="1"/>
    <brk id="9" max="1048575" man="1"/>
</colBreaks>
```

---

## Background Image

### In sheet.xml
```xml
<picture r:id="rId2"/>
```

### In sheet.xml.rels
```xml
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.png"/>
```

---

## Print Titles (rows and columns repeated on each page)

### In workbook.xml
```xml
<definedName name="_xlnm.Print_Titles" localSheetId="0">Sheet1!$A:$B,Sheet1!$1:$2</definedName>
```

---

## Headers and Footers

### Examples
```xml
<headerFooter>
    <oddHeader>Page &amp;P</oddHeader>
    <oddFooter>Page &amp;P</oddFooter>
</headerFooter>

<headerFooter>
    <oddHeader>Page &amp;P of &amp;N</oddHeader>
    <oddFooter>&amp;A</oddFooter>
</headerFooter>

<headerFooter differentOddEven="1" differentFirst="1">
    <oddHeader>&amp;L&amp;B Confidential&amp;B&amp;C&amp;D&amp;RPage &amp;P</oddHeader>
    <oddFooter>&amp;F</oddFooter>
</headerFooter>
```

### Header/Footer Codes
| Code | Meaning |
|------|---------|
| &P | Page number |
| &N | Page count |
| &A | Sheet name |
| &D | Date |
| &F | Filename |
| &Z | Path |
| &L | Align Left |
| &C | Align Centre |
| &R | Align Right |
| &B | Bold Toggle |
| &"Font Name" | Font |
| &14 | Font size |
| &S | Strikethrough |
| &U | Underline |
| &X | Superscript |
| &KFF0000 | RGB color |

---

## Page Setup API

```javascript
sheet.pageSetup = {
    margins: {
        left: 0.25, right: 0.25, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3
    },
    orientation: 'portrait' | 'landscape',
    pageOrder: 'overThenDown' | 'downThenOver',
    blackAndWhite: boolean,
    draft: boolean,
    cellComments: 'atEnd' | 'asDisplayed' | 'None',
    errors: 'dash' | 'blank' | 'NA' | 'displayed',
    scale: 90,
    fitToWidth: 2,
    fitToHeight: 3,
    paperSize: 9,
    showRowColHeaders: boolean,
    showGridLines: boolean,
    horizontalCentered: boolean,
    verticalCentered: boolean,
    rowBreaks: [5, 7],
    colBreaks: [3, 6, 9],
    
    // headers/footers
    header: {
        odd: { left: '', center: '&[Page] of &[Pages]', right: '' },
        even: {},
        first: {}
    },
    footer: {},
    
    // defined names
    printArea: 'A1:D11',
    printTitles: {
        rows: '1:2',
        columns: 'B:C'
    }
}
```

---

## Phonetic Text (Japanese)

```xml
<si>
    <t>役割</t>  <!-- "role" in kanji -->
    <rPh sb="0" eb="2">
        <t>ヤクワリ</t> <!-- pronunciation in katakana -->
    </rPh>
    <phoneticPr fontId="1" />
</si>
```

---

## Comments

```xml
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <authors>
        <author>Author</author>
    </authors>
    <commentList>
        <comment ref="A1" authorId="0" shapeId="0">
            <text>
                <r>
                    <rPr>
                        <b/><sz val="9"/><color indexed="81"/>
                        <rFont val="Tahoma"/><charset val="1"/>
                    </rPr>
                    <t>Author:</t>
                </r>
                <r>
                    <rPr>
                        <sz val="9"/><color indexed="81"/>
                        <rFont val="Tahoma"/><charset val="1"/>
                    </rPr>
                    <t xml:space="preserve">
Is it me you're looking for</t>
                </r>
            </text>
        </comment>
    </commentList>
</comments>
```

---

## Charts

### New Contents for Chart as Sheet

**[Content_Types].xml**: chartsheets, charts, drawings, etc

**app.xml**:
- HeadingPairs: Charts: 1
- TitleOfParts: Chart sheet

**File locations**:
- `/xl/charts/*`
- `/xl/chartsheets/*`
- `/xl/drawings/`

---

## Conditional Formatting

### Rule Types
- `expression`: formula returns true/false
- `cellIs`: operator + formulae
- `top10`: percent, bottom, rank
- `aboveAverage`: aboveAverage flag
- `dataBar`: cfvo, color, showValue, gradient, border, etc.
- `colorScale`: cfvo (2-3), colors
- `iconSet`: iconSet enum, cfvo (2-5), reverse, showValue, custom
- `containsText`: operator, text, formulae
- `timePeriod`: thisWeek, lastWeek, nextWeek, yesterday, today, tomorrow, etc.

### CFVO (Conditional Formatting Value Object)
```javascript
{
    type: 'min' | 'max' | 'percentile' | 'num' | 'percent' | 'formula' | 'autoMin' | 'autoMax',
    val: number
}
```

### Icon Sets
- Native: 3Arrows, 4Arrows, 5Arrows, 3ArrowsGray, 4ArrowsGray, 5ArrowsGray, 3TrafficLights, 3Signs, 4RedToBlack, 3TrafficLights2, 4TrafficLights, 3Symbols3, 3Flags, Symbols2, 5Quarters, 4Rating, 5Rating
- Requires extLst: 3Triangles, 3Stars, 5Boxes

### Data Bar Options
```javascript
{
    cfvo: [CFVO, ...],
    color: Colour,
    showValue: true | false,
    gradient: true | false,
    border: true | false,
    borderColour: Colour,
    axisColor: Colour,
    negativeFillColor: Colour,
    negativeBorderColor: Colour,
    negativeBarColorSameAsPositive: boolean,
    negativeBarBorderColorSameAsPositive: boolean,
    minLength: 0..maxLength,
    maxLength: minLength..100,
    axisPosition: 'middle' | 'none' | undefined,
    direction: 'rightToLeft' | undefined
}
```
