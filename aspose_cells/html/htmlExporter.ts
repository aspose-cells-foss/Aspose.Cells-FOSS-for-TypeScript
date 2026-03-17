import { Workbook } from "../workbook";
import { Worksheet } from "../worksheet";
import { join } from "path";
import type { ShapeInfo, ChartInfo } from "../types";
import { ChartRenderer } from "./chartRenderer";

const EMU_PER_PIXEL = 914400 / 96;

export class HtmlExporter {
  private workbook: Workbook;
  private chartRenderer: ChartRenderer;

  constructor(workbook: Workbook) {
    this.workbook = workbook;
    this.chartRenderer = new ChartRenderer();
  }

  async saveAsHtmlFrameset(
    filePath: string,
    htmlDir: string,
    fileName: string,
  ) {
    const { writeFile, mkdir } = await import("fs/promises");
    const { existsSync } = await import("fs");

    if (!existsSync(htmlDir)) {
      await mkdir(htmlDir, { recursive: true });
    }

    const mainHtml = this.generateFramesetHtml(fileName + "_files");
    await writeFile(filePath, mainHtml);

    const filelistXml = this.generateFilelistXml(fileName);
    await writeFile(join(htmlDir, "filelist.xml"), filelistXml);

    const stylesheetCss = this.generateStylesheetCss();
    await writeFile(join(htmlDir, "stylesheet.css"), stylesheetCss);

    const tabstripHtml = this.generateTabstripHtml(
      fileName + "_files",
      this.workbook.worksheets.map((w: Worksheet) => w.name),
      fileName,
      this.workbook.activeTab,
    );
    await writeFile(join(htmlDir, "tabstrip.htm"), tabstripHtml);

    for (let i = 0; i < this.workbook.worksheets.length; i++) {
      const sheet = this.workbook.worksheets.get(i)!;
      const sheetHtml = this.generateSheetHtml(sheet, fileName, i === 0, i);
      await writeFile(
        join(htmlDir, `sheet${String(i + 1).padStart(3, "0")}.htm`),
        sheetHtml,
      );
    }

    await writeFile(join(htmlDir, "editdata.mso"), "");
    await writeFile(join(htmlDir, "oledata.mso"), "");
  }

  private generateFramesetHtml(filesDir: string): string {
    const sheetRefs = this.workbook.worksheets
      .map(
        (w, i) =>
          `    <x:ExcelWorksheet>\n     <x:Name>${this.escapeHtml(w.name)}</x:Name>\n     <x:WorksheetSource HRef="${filesDir}/sheet${String(i + 1).padStart(3, "0")}.htm"/>\n    </x:ExcelWorksheet>`,
      )
      .join("\n");

    const tabLinks = this.workbook.worksheets
      .map(
        (_, i) =>
          `<link id="shLink" href="${filesDir}/sheet${String(i + 1).padStart(3, "0")}.htm"/>`,
      )
      .join("\n");

    const tabNames = this.workbook.worksheets
      .map((w, i) => `c_rgszSh[${i}]="${w.name}";`)
      .join("\n  ");

    return `﻿<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Frameset//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-frameset.dtd">
<html xmlns:v="urn:schemas-microsoft-com:vml"
xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta name="Excel Workbook Frameset" content=""/>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
<meta name="ProgId" content="Excel.Sheet"/>
<meta name="Generator" content="Aspose.Cells FOSS for TypeScript"/>
<link rel="File-List" href="${filesDir}/filelist.xml"/>
<link rel="Edit-Time-Data" href="${filesDir}/editdata.mso"/>
<link rel="OLE-Object-Data" href="${filesDir}/oledata.mso"/>
<![if !supportTabStrip]>
${tabLinks}

<link id="shLink"/>

<script language="JavaScript" type="text/javascript">
<!--
 var c_lTabs=${this.workbook.worksheets.length};

 var c_rgszSh=new Array(c_lTabs);
 ${tabNames}

 var c_rgszClr=new Array(8);
 c_rgszClr[0]="window";
 c_rgszClr[1]="buttonface";
 c_rgszClr[2]="windowframe";
 c_rgszClr[3]="windowtext";
 c_rgszClr[4]="threedlightshadow";
 c_rgszClr[5]="threedhighlight";
 c_rgszClr[6]="threeddarkshadow";
 c_rgszClr[7]="threedshadow";

 var g_iShCur;
 var g_rglTabX=new Array(c_lTabs);

function fnTabToCol(iTab)
{
 return 2*iTab+1;
}

function fnNextTab(fDir)
{
 var iNextTab=-1;
 var i;

 with (frames['frTabs'].document.body) {
  if (fDir==0) {
   if (scrollLeft>0) {
    for (i=0;i<c_lTabs&&g_rglTabX[i]<scrollLeft;i++);
    if (i<c_lTabs)
     iNextTab=i-1;
   }
  } else {
   if (g_rglTabX[c_lTabs]+6>offsetWidth+scrollLeft) {
    for (i=0;i<c_lTabs&&g_rglTabX[i]<=scrollLeft;i++);
    if (i<c_lTabs)
     iNextTab=i;
   }
  }
 }
 return iNextTab;
}

function fnScrollTabs(fDir)
{
 var iNextTab=fnNextTab(fDir);

 if (iNextTab>=0) {
  frames['frTabs'].scroll(g_rglTabX[iNextTab],0);
  return true;
 } else
  return false;
}

function fnFastScrollTabs(fDir)
{
 if (c_lTabs>16)
  frames['frTabs'].scroll(g_rglTabX[fDir?c_lTabs-1:0],0);
 else
  if (fnScrollTabs(fDir)>0) window.setTimeout("fnFastScrollTabs("+fDir+");",5);
}

function fnSetTabProps(iTab,fActive)
{
 var iCol=fnTabToCol(iTab);
 var i;

 if (iTab>=0) {
  with (frames['frTabs'].document.body) {
   with (tbTabs) {
    for (i=0;i<=4;i++) {
     with (rows[i]) {
      if (i==0)
       cells[iCol].style.background=c_rgszClr[fActive?0:2];
      else if (i>0 && i<4) {
       if (fActive) {
        cells[iCol-1].style.background=c_rgszClr[2];
        cells[iCol].style.background=c_rgszClr[0];
        cells[iCol+1].style.background=c_rgszClr[2];
       } else {
        if (i==1) {
         cells[iCol-1].style.background=c_rgszClr[2];
         cells[iCol].style.background=c_rgszClr[1];
         cells[iCol+1].style.background=c_rgszClr[2];
        } else {
         cells[iCol-1].style.background=c_rgszClr[4];
         cells[iCol].style.background=c_rgszClr[(i==2)?2:4];
         cells[iCol+1].style.background=c_rgszClr[4];
        }
       }
      } else
       cells[iCol].style.background=c_rgszClr[fActive?2:4];
     }
    }
   }
   with (aTab[iTab].style) {
    cursor=(fActive?"default":"hand");
    color=c_rgszClr[3];
   }
  }
 }
}

function fnMouseOverScroll(iCtl)
{
 frames['frScroll'].document.all.tdScroll[iCtl].style.color=c_rgszClr[7];
}

function fnMouseOutScroll(iCtl)
{
 frames['frScroll'].document.all.tdScroll[iCtl].style.color=c_rgszClr[6];
}

function fnTab)
{
 ifMouseOverTab(i (iTab!=g_iShCur) {
  var iCol=fnTabToCol(iTab);
  with (frames['frTabs'].document.all) {
   tdTab[iTab].style.background=c_rgszClr[5];
  }
 }
}

function fnMouseOutTab(iTab)
{
 if (iTab>=0) {
  var elFrom=frames['frTabs'].event.srcElement;
  var elTo=frames['frTabs'].event.toElement;

  if ((!elTo) ||
   (elFrom.tagName==elTo.tagName) ||
   (elTo.tagName=="A" && elTo.parentElement!=elFrom) ||
   (elFrom.tagName=="A" && elFrom.parentElement!=elTo)) {

   if (iTab!=g_iShCur) {
    with (frames['frTabs'].document.all) {
     tdTab[iTab].style.background=c_rgszClr[1];
    }
   }
  }
 }
}

function fnSetActiveSheet(iSh)
{
 if (iSh!=g_iShCur) {
  fnSetTabProps(g_iShCur,false);
  fnSetTabProps(iSh,true);
  g_iShCur=iSh;
 }
}

//-->
</script>
<![endif]-->
<!--[if gte mso 9]><xml>
 <x:ExcelWorkbook>
  <x:ExcelWorksheets>
 ${sheetRefs}
  </x:ExcelWorksheets>
  <x:Stylesheet HRef="${filesDir}/stylesheet.css"/>
  <x:WindowHeight>15500</x:WindowHeight>
  <x:WindowWidth>25820</x:WindowWidth>
  <x:WindowTopX>-110</x:WindowTopX>
  <x:WindowTopY>-110</x:WindowTopY>
  <x:RefModeR1C1/>
  <x:TabRatio>600</x:TabRatio>
   <x:ActiveSheet>0</x:ActiveSheet>
  </x:ExcelWorkbook>
</xml><![endif]-->
</head>
    <frameset rows="*,39">
     <frame src="${filesDir}/sheet001.htm" name="frSheet"/>
     <frame src="${filesDir}/tabstrip.htm" name="frTabs" marginwidth="0" marginheight="0"/>
     <noframes>
      <body>
       <p>This page uses frames, but your browser doesn't support them.</p>
      </body>
     </noframes>
    </frameset>
</html>`;
  }

  private generateFilelistXml(mainFileName: string): string {
    const files = [
      `../${mainFileName}.html`,
      "editdata.mso",
      "stylesheet.css",
      "tabstrip.htm",
      ...this.workbook.worksheets.map(
        (_, i) => `sheet${String(i + 1).padStart(3, "0")}.htm`,
      ),
      "filelist.xml",
    ];

    return (
      `﻿<xml xmlns:o="urn:schemas-microsoft-com:office:office">\n` +
      files.map((f) => ` <o:File HRef="${f}"/>`).join("\n") +
      `\n</xml>`
    );
  }

  private generateStylesheetCss(): string {
    const styles = this.workbook.worksheets.styles;
    let css = `﻿tr
 {mso-height-source:auto;
 mso-ruby-visibility:none;}
col
 {mso-width-source:auto;
 mso-ruby-visibility:none;}
br
 {mso-data-placement:same-cell;}
ruby
 {ruby-align:start;}
.style0
 {
 mso-number-format:General;
 text-align:general;
 vertical-align:bottom;
 white-space:nowrap;
 background:auto;
 mso-pattern:auto;
 color:#000000;
 font-size:10pt;
 font-weight:400;
 font-style:normal;
 font-family:Arial,sans-serif;
 mso-protection:locked visible;
 mso-style-name:Normal;
 mso-style-id:0;}
td
 {mso-style-parent:style0;
 mso-number-format:General;
 text-align:general;
 vertical-align:bottom;
 white-space:nowrap;
 background:auto;
 mso-pattern:auto;
 color:#000000;
 font-size:10pt;
 font-weight:400;
 font-style:normal;
 font-family:Arial,sans-serif;
 mso-protection:locked visible;
 mso-ignore:padding;
 }
`;

    for (const [index, style] of styles) {
      css += this.generateStyleCss(index + 1, style);
    }

    return css;
  }

  private generateStyleCss(index: number, style: any): string {
    const numFmt = this.escapeCss(
      (style.numberFormat || "General").replace(/"/g, '\\"'),
    );
    const align = style.alignment?.horizontal || "general";
    const valign = style.alignment?.vertical || "bottom";
    const wrapText = style.alignment?.wrapText
      ? "normal;word-wrap:break-word"
      : "nowrap";

    let bgColor = "auto";
    let msoPattern = "auto";
    if (style.fill?.fgColor) {
      bgColor = this.convertColor(style.fill.fgColor);
      msoPattern = style.fill.patternType === "solid" ? "solid" : "auto";
    } else if (style.fill?.bgColor) {
      bgColor = this.convertColor(style.fill.bgColor);
      msoPattern = style.fill.patternType === "solid" ? "solid" : "auto";
    }

    const fontSize = style.font?.size || 10;
    const fontWeight = style.font?.bold ? "bold" : "400";
    const fontStyle = style.font?.italic ? "italic" : "normal";
    const fontFamily = style.font?.name || "Arial";
    const fontColor = this.convertColor(style.font?.color) || "#000000";

    let borderCss = "";
    const borderMap: [string, any][] = [
      ["top", style.border?.top],
      ["right", style.border?.right],
      ["bottom", style.border?.bottom],
      ["left", style.border?.left],
    ];
    for (const [side, border] of borderMap) {
      if (border?.style) {
        const borderStyle =
          border.style === "thick"
            ? "thick"
            : border.style === "medium"
              ? "medium"
              : "thin";
        const borderColor = this.convertColor(border.color) || "windowtext";
        borderCss += `\n border-${side}:1px solid ${borderColor};`;
      } else {
        borderCss += `\n border-${side}:none;`;
      }
    }

    return `.x${index}
 {
 mso-number-format:"${numFmt}";
 text-align:${align};
 vertical-align:${valign};
 white-space:${wrapText};
 background:${bgColor};
 mso-pattern:${msoPattern};
 color:${fontColor};
 font-size:${fontSize}pt;
 font-weight:${fontWeight};
 font-style:${fontStyle};
 font-family:${fontFamily},sans-serif;${borderCss}
 mso-diagonal-down:none;
 mso-diagonal-up:none;
 mso-protection:locked visible;
 }
`;
  }

  private convertColor(color: string | undefined): string {
    if (!color) return "";
    if (color.startsWith("#")) {
      return color;
    }
    if (/^[0-9A-Fa-f]{8}$/.test(color)) {
      return "#" + color.substring(2);
    }
    if (/^[0-9A-Fa-f]{6}$/.test(color)) {
      return "#" + color;
    }
    if (/^[0-9A-Fa-f]{3}$/.test(color)) {
      return (
        "#" + color[0] + color[0] + color[1] + color[1] + color[2] + color[2]
      );
    }
    return "";
  }

  private escapeCss(str: string): string {
    return str.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
  }

  private generateTabstripHtml(
    filesDir: string,
    sheetNames: string[],
    mainFileName: string,
    activeTab: number = 0,
  ): string {
    const tabs = sheetNames
      .map((name, i) => {
        const isActive = i === activeTab ? "active" : "";
        return `<td nowrap='nowrap' class='tab ${isActive}'><b><small><small>&nbsp;<a href="sheet${String(i + 1).padStart(3, "0")}.htm" target="frSheet"><font face="Arial" color="#000000">${name}</font></a>&nbsp;</small></small></b></td>`;
      })
      .join(" ");

    return `﻿<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
<meta name="ProgId" content="Excel.Sheet"/>
<meta name="Generator" content="Aspose.Cells FOSS for TypeScript"/>
<link id="Main-File" rel="Main-File" href="../${mainFileName}.html">

<script language="JavaScript">
<!--
if (window.name!="frTabs")
 window.location.replace(document.all.item("Main-File").href);
window.addEventListener('load', function(){ document.querySelectorAll('.tab').forEach(tab => {
  tab.addEventListener('click', function() {
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    this.classList.add('active');
 });});});
//-->
</script>
<style>
<!--
A {
    text-decoration:none;
    color:#000000;
    font-size:9pt;
}
-->
.tab { 
  background: #ddd;
}
.tab.active{
  background: #fff;
}
</style>
</head>
<body topmargin='0' leftmargin='0' bgcolor="#808080">
<table border='0' cellspacing='1'>
 <tr> ${tabs}

 </tr>
</table>
</body>
</html>`;
  }

  private generateSheetHtml(
    worksheet: Worksheet,
    mainFileName: string,
    isFirst: boolean,
    sheetIndex: number,
  ): string {
    const tableHtml = this.worksheetToHtmlForExcel(worksheet);

    return `﻿<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns:v="urn:schemas-microsoft-com:vml"
xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
<meta name="ProgId" content="Excel.Sheet"/>
<meta name="Generator" content="Aspose.Cells FOSS for TypeScript"/>
<link id="Main-File" rel="Main-File" href="../${mainFileName}.html"/>
<link rel="File-List" href="filelist.xml"/>
<link rel="Edit-Time-Data" href="editdata.mso"/>
<!--[if !mso]>
<style  type="text/css">
v\:* {behavior:url(#default#VML);}
o\:* {behavior:url(#default#VML);}
x\:* {behavior:url(#default#VML);}
.shape {behavior:url(#default#VML);}
</style>
<![endif]-->
<link rel="Stylesheet" href="stylesheet.css"/>
<style type="text/css">
<!--table
 {mso-displayed-decimal-separator:"\.";
 mso-displayed-thousand-separator:",";}
@page
 {
 mso-header-data:"";
 mso-footer-data:"";
 margin:1in 0.75in 1in 0.75in;
 mso-header-margin:0.5in;
 mso-footer-margin:0.5in;
 mso-page-orientation:Portrait;
 }
-->
</style>
<![if !supportTabStrip]>
<script language="JavaScript" type="text/javascript">
<!--

if (window.name!="frSheet")
 window.location.replace("../${mainFileName}.html");
//-->
</script>
<![endif]-->
<!--[if gte mso 9]><xml>
<x:WorksheetOptions>
 <x:StandardWidth>2048</x:StandardWidth>
 <x:Print>
  <x:ValidPrinterInfo/>
  <x:PaperSizeIndex>1</x:PaperSizeIndex>
  <x:HorizontalResolution>600</x:HorizontalResolution>
  <x:VerticalResolution>600</x:VerticalResolution>
 </x:Print>
 ${!isFirst ? "" : "<x:Selected/>"}
</x:WorksheetOptions>
</xml><![endif]-->
</head>

<body link='blue' vlink='purple' >

${tableHtml}

</body>


</html>`;
  }

  worksheetToHtmlForExcel(worksheet: Worksheet): string {
    const cells = Array.from(worksheet.cells);
    if (cells.length === 0) {
      return `<table border='0' cellpadding='0' cellspacing='0' style='border-collapse:collapse;table-layout:fixed;'></table>`;
    }

    const shapes = worksheet.shapes;
    const maxRow = Math.max(
      ...cells.map((c) => c.row),
      ...shapes.map((s) => s.toRow),
    );
    const maxCol = Math.max(
      ...cells.map((c) => c.col),
      ...shapes.map((s) => s.toCol),
    );

    const defaultWidth = 64;
    const columnWidths: number[] = [];
    for (let col = 0; col <= maxCol; col++) {
      columnWidths.push(worksheet.getColumnWidth(col) || defaultWidth);
    }

    const rowHeights: number[] = [];
    for (let row = 0; row <= maxRow; row++) {
      rowHeights.push(worksheet.getRowHeight(row) || 16);
    }

    let totalWidth = 0;
    let colWidthsStr = "";
    let span = 1;
    for (let col = 0; col <= maxCol; col++) {
      const width = columnWidths[col];
      totalWidth += width;

      const nextCol = col + 1;
      if (nextCol <= maxCol && columnWidths[nextCol] === width) {
        span++;
      } else {
        const ptWidth = width * 0.75;
        colWidthsStr += `<col width='${width}' span='${span}' style='mso-width-source:userset;width:${ptWidth}pt'/>`;
        span = 1;
      }
    }

    let tableHtml = `<table border='0' cellpadding='0' cellspacing='0' width='${totalWidth}' style='border-collapse:collapse;table-layout:fixed;width:${totalWidth * 0.75}pt'>
 ${colWidthsStr}
`;

    const usedStyles = new Set<number>();
    cells.forEach((c) => {
      if (c.styleIndex !== undefined) usedStyles.add(c.styleIndex);
    });

    for (let row = 0; row <= maxRow; row++) {
      const rowHeight = worksheet.getRowHeight(row) || 16;
      tableHtml += ` <tr height='${rowHeight}' style='mso-height-source:userset;height:${rowHeight * 0.75}pt'>`;

      for (let col = 0; col <= maxCol; col++) {
        const cell = worksheet.getCell(row, col);
        const rawValue = cell?.value;
        const value =
          rawValue !== null && rawValue !== undefined ? String(rawValue) : "";

        const isNumber = typeof rawValue === "number";
        const styleIndex = cell?.styleIndex;
        const colWidth = columnWidths[col];

        let cellAttr = "";
        if (styleIndex !== undefined) {
          cellAttr = `class='x${styleIndex + 1}'`;
        }

        if (col === 0) {
          cellAttr += ` height='${rowHeight}'`;
          if (row === 0) {
            cellAttr += ` width='${colWidth}' style='height:${rowHeight * 0.75}pt;width:${colWidth * 0.75}pt;'`;
          } else {
            cellAttr += ` style='height:${rowHeight * 0.75}pt;'`;
          }
        } else if (row === 0) {
          cellAttr += ` width='${colWidth}' style='width:${colWidth * 0.75}pt;'`;
        }

        // Find shapes that start at this cell
        const cellShapes = shapes.filter(
          (s) => s.fromRow === row && s.fromCol === col,
        );

        if (cellShapes.length > 0) {
          let shapesInCell = "";
          for (const shape of cellShapes) {
            shapesInCell += this.shapeToCellSvg(
              shape,
              columnWidths,
              rowHeights,
              worksheet,
            );
          }

          const cellContent = this.escapeHtml(value);
          tableHtml += `\n<td ${cellAttr} valign='top' align='left'>${cellContent}${shapesInCell}</td>`;
        } else if (isNumber) {
          tableHtml += `\n<td ${cellAttr} align='right'>${this.escapeHtml(value)}</td>`;
        } else {
          tableHtml += `\n<td ${cellAttr}>${this.escapeHtml(value)}</td>`;
        }
      }

      tableHtml += "\n </tr>\n";
    }

    tableHtml += `<![if supportMisalignedColumns]>
  <tr height='0' style='display:none'>
   <td width='${totalWidth}' colspan='${maxCol + 1}' style='width:${totalWidth * 0.75}pt;mso-ignore:colspan;'></td>
  </tr>
  <![endif]>
</table>`;

    return tableHtml;
  }

  private shapeToCellSvg(
    shape: ShapeInfo,
    columnWidths: number[],
    rowHeights: number[],
    worksheet: Worksheet,
  ): string {
    if ("chartType" in shape) {
      return this.chartRenderer.renderChart(shape as ChartInfo, worksheet);
    }

    let classCounter = 0;
    const getClassName = () =>
      `p${shape.name.replace(/\s/g, "_")}_${classCounter++}`;

    const fromColOff = shape.fromColOff ?? 0;
    const fromRowOff = shape.fromRowOff ?? 0;
    const toColOff = shape.toColOff ?? 0;
    const toRowOff = shape.toRowOff ?? 0;

    let left = fromColOff / EMU_PER_PIXEL;
    let top = fromRowOff / EMU_PER_PIXEL;

    let right = 0;
    for (let col = shape.fromCol; col < shape.toCol; col++) {
      right += columnWidths[col] || 64;
    }
    right += toColOff / EMU_PER_PIXEL;

    let bottom = 0;
    for (let row = shape.fromRow; row < shape.toRow; row++) {
      bottom += rowHeights[row] || 16;
    }
    bottom += toRowOff / EMU_PER_PIXEL;

    const width = right - left;
    const height = bottom - top;

    const schemeColorMap: Record<string, string> = {
      accent1: "#4472C4",
      accent2: "#ED7D31",
      accent3: "#A5A5A5",
      accent4: "#FFC000",
      accent5: "#5B9BD5",
      accent6: "#70AD47",
    };

    let fillColor = "#ffffff";
    let hasFill = false;
    let gradientDefs = "";
    if (shape.fill && shape.fill.type !== "none") {
      hasFill = true;
      if (shape.fill.type === "solid" && shape.fill.color) {
        if (shape.fill.isSchemeColor && schemeColorMap[shape.fill.color]) {
          fillColor = schemeColorMap[shape.fill.color];
        } else if (shape.fill.color.startsWith("#")) {
          fillColor = shape.fill.color.toUpperCase();
        } else if (schemeColorMap[shape.fill.color]) {
          fillColor = schemeColorMap[shape.fill.color];
        } else {
          fillColor = shape.fill.color;
        }
      } else if (shape.fill.type === "gradient" && shape.fill.gradientStops) {
        const gradientId = `grad_${shape.name.replace(/\s/g, "_")}`;
        const angle = shape.fill.gradientAngle
          ? shape.fill.gradientAngle / 10000
          : 0;
        let gradientStopsXml = "";
        for (const stop of shape.fill.gradientStops) {
          let stopColor = stop.color;
          if (stop.isSchemeColor && schemeColorMap[stop.color]) {
            stopColor = schemeColorMap[stop.color];
          } else if (!stop.color.startsWith("#")) {
            stopColor = schemeColorMap[stop.color] || stop.color;
          }
          gradientStopsXml += `<stop offset="${stop.position / 1000}" stop-color="${stopColor.toUpperCase()}" stop-opacity="1" />`;
        }
        gradientDefs = `
   <linearGradient id="${gradientId}" x1="0%" x2="0%" y1="0%" y2="100%" gradientTransform="rotate(${angle - 90}, 0.5, 0.5)">
    ${gradientStopsXml}
   </linearGradient>`;
        fillColor = `url(#${gradientId})`;
      }
    } else if (
      shape.fillColor &&
      shape.fillColor !== "#ffffff" &&
      shape.fillColor !== "white"
    ) {
      hasFill = true;
      fillColor = shape.fillColor.toUpperCase();
    }

    const strokeColor = shape.lineColor || "#000000";

    const type = shape.type.toLowerCase();
    const isLine = type === "line" || type.includes("connector");
    const isBentConnector =
      type.includes("bent") ||
      type.includes("elbow") ||
      type.includes("curved");

    const svgWidth = (width / 1.33333 + 0.75).toFixed(2) + "pt";
    const svgHeight = (height / 1.33333 + 0.75).toFixed(2) + "pt";

    let fillPath = "";
    let strokePath = "";
    let transform = "";

    const scale = 1.33333;
    const px = 0.5;
    const py = 0.5;
    const pw = width / scale - 0.5;
    const ph = height / scale - 0.5;

    const rotationRad = ((shape.rotation || 0) * Math.PI) / 180;
    const cosR = Math.cos(rotationRad);
    const sinR = Math.sin(rotationRad);

    if (shape.rotation && shape.rotation !== 0) {
      const cx = width / scale / 2;
      const cy = height / scale / 2;
      transform = `matrix(${cosR},${sinR},${-sinR},${cosR},${cy * sinR + cx - cx * cosR},${cy - cx * sinR - cy * cosR})`;
    } else if (shape.flipV) {
      transform = `matrix(1,0,0,-1,0,${height / EMU_PER_PIXEL})`;
    } else if (shape.flipH) {
      transform = `matrix(-1,0,0,1,${width / EMU_PER_PIXEL},0)`;
    }

    if (isLine) {
      const className = getClassName();

      if (isBentConnector) {
        let cornerX = 0;
        for (let col = shape.fromCol; col < shape.toCol; col++) {
          cornerX += columnWidths[col] || 64;
        }
        cornerX = cornerX / scale;
        const cornerY = fromRowOff / EMU_PER_PIXEL / scale;

        const endY = bottom / scale;
        const endX = right / scale;
        const startX = fromColOff / EMU_PER_PIXEL / scale;

        strokePath = `<path d="M${startX.toFixed(6)},${cornerY.toFixed(6)} L${cornerX.toFixed(6)},${cornerY.toFixed(6)} L${cornerX.toFixed(6)},${endY.toFixed(6)} L${endX.toFixed(6)},${endY.toFixed(6)} " class="${className}" fill="none" transform="matrix(1,0,0,1,-0.000007629,-0.000007629)" />`;
      } else {
        strokePath = `<path d="M${left / scale},${top / scale} L${right / scale},${bottom / scale} " class="${className}" fill="none" transform="matrix(1,0,0,1,-0.000007629,-0.000007629)" />`;
      }

      if (shape.hasArrowEnd) {
        const dx = right - left;
        const dy = bottom - top;
        const angle = Math.atan2(dy, dx);
        const arrowSize = 8;
        const arrowAngle = Math.PI / 6;

        const ax1 = right - arrowSize * Math.cos(angle - arrowAngle);
        const ay1 = bottom - arrowSize * Math.sin(angle - arrowAngle);
        const ax2 = right - arrowSize * Math.cos(angle + arrowAngle);
        const ay2 = bottom - arrowSize * Math.sin(angle + arrowAngle);

        const arrowPath = `M${right / scale},${bottom / scale} L${ax1 / scale},${ay1 / scale} L${ax2 / scale},${ay2 / scale} Z`;
        fillPath = `<path d="${arrowPath}" fill="${strokeColor}" transform="matrix(1,0,0,1,-0.000007629,-0.000007629)" />`;
      }
    } else if (type === "ellipse" || type === "oval") {
      const cx = px + pw / 2;
      const cy = py + ph / 2;
      const rx = pw / 2;
      const ry = ph / 2;
      const ellipsePath = `M${cx},${cy - ry} C${cx + rx},${cy - ry} ${cx + rx},${cy + ry} ${cx},${cy + ry} C${cx - rx},${cy + ry} ${cx - rx},${cy - ry} ${cx},${cy - ry} Z`;
      fillPath = hasFill
        ? `<path d="${ellipsePath}" fill="${fillColor}"${transform ? ` transform="${transform}"` : ""} />`
        : "";
      const className = getClassName();
      strokePath = `<path d="${ellipsePath}" class="${className}" fill="none"${transform ? ` transform="${transform}"` : ""} />`;
    } else if (type === "triangle" || type === "rightTriangle") {
      let pathD: string;
      if (type === "rightTriangle") {
        pathD = `M${px},${py + ph} L${px},${py} L${px + pw},${py + ph} Z`;
      } else {
        pathD = `M${px + pw / 2},${py} L${px},${py + ph} L${px + pw},${py + ph} Z`;
      }
      fillPath = hasFill
        ? `<path d="${pathD} " fill="${fillColor}"${transform ? ` transform="${transform}"` : ""} />`
        : "";
      const className = getClassName();
      strokePath = `<path d="${pathD} " class="${className}" fill="none"${transform ? ` transform="${transform}"` : ""} />`;
    } else if (type === "diamond") {
      const pathD = `M${px + pw / 2},${py} L${px + pw},${py + ph / 2} L${px + pw / 2},${py + ph} L${px},${py + ph / 2} Z`;
      fillPath = hasFill
        ? `<path d="${pathD} " fill="${fillColor}"${transform ? ` transform="${transform}"` : ""} />`
        : "";
      const className = getClassName();
      strokePath = `<path d="${pathD} " class="${className}" fill="none"${transform ? ` transform="${transform}"` : ""} />`;
    } else {
      const pathD = `M${px},${py} L${px + pw},${py} L${px + pw},${py + ph} L${px},${py + ph} Z`;
      fillPath = hasFill
        ? `<path d="${pathD} " fill="${fillColor}"${transform ? ` transform="${transform}"` : ""} />`
        : "";
      const className = getClassName();
      strokePath = `<path d="${pathD} " class="${className}" fill="none"${transform ? ` transform="${transform}"` : ""} />`;
    }

    const strokeClassName = strokePath.match(/class="([^"]+)"/)?.[1] || "";
    const styleBlock = strokeClassName
      ? `
   <style type="text/css">
    <![CDATA[
    
    
.${strokeClassName}
{
stroke:${strokeColor};
stroke-width:1px;
stroke-linecap:butt;
stroke-linejoin:miter;
}

    ]]></style>`
      : "";

    const svgContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" width="${svgWidth}" height="${svgHeight}">
  <defs>
    ${gradientDefs}
    ${styleBlock}
  </defs>
  <g id="SFixTitle" />
  <g id="SContent">
   <g transform="scale(1.33333)">
    <rect width="100%" height="100%" fill="white" />
    <g>
     <g>
      <g>${
        fillPath
          ? `
       ${fillPath}`
          : ""
      }
       <g />
${strokePath ? `      ${strokePath}` : ""}
     </g>
     </g>
    </g>
   </g>
</svg>`;

    const leftPx = left / EMU_PER_PIXEL;
    const topPx = top / EMU_PER_PIXEL;
    return `<span style='mso-ignore:vglayout;position:absolute;z-index:1;margin-left:${left}px;margin-top:${top}px;width:${width}px;height:${height}px'>${svgContent}</span>`;
  }

  private generateShapesHtml(worksheet: Worksheet): string {
    const shapes = worksheet.shapes;
    if (shapes.length === 0) return "";

    const cells = Array.from(worksheet.cells);
    const maxRow = cells.length > 0 ? Math.max(...cells.map((c) => c.row)) : 0;
    const maxCol = cells.length > 0 ? Math.max(...cells.map((c) => c.col)) : 0;

    const columnWidths: number[] = [];
    for (let col = 0; col <= maxCol; col++) {
      columnWidths.push(worksheet.getColumnWidth(col) || 64);
    }

    const rowHeights: number[] = [];
    for (let row = 0; row <= maxRow; row++) {
      rowHeights.push(worksheet.getRowHeight(row) || 16);
    }

    // Calculate bounds from shapes
    let maxX = 0;
    let maxY = 0;

    const processedShapes = shapes.map((shape) => {
      const fromColOff = shape.fromColOff ?? 0;
      const fromRowOff = shape.fromRowOff ?? 0;
      const toColOff = shape.toColOff ?? 0;
      const toRowOff = shape.toRowOff ?? 0;

      let left = 0;
      for (let col = 0; col < shape.fromCol; col++) {
        left += columnWidths[col] || 64;
      }
      left += fromColOff / EMU_PER_PIXEL;

      let top = 0;
      for (let row = 0; row < shape.fromRow; row++) {
        top += rowHeights[row] || 16;
      }
      top += fromRowOff / EMU_PER_PIXEL;

      let right = 0;
      for (let col = 0; col < shape.toCol; col++) {
        right += columnWidths[col] || 64;
      }
      right += toColOff / EMU_PER_PIXEL;

      let bottom = 0;
      for (let row = 0; row < shape.toRow; row++) {
        bottom += rowHeights[row] || 16;
      }
      bottom += toRowOff / EMU_PER_PIXEL;

      if (right > maxX) maxX = right;
      if (bottom > maxY) maxY = bottom;

      return { shape, left, top, right, bottom };
    });

    const totalWidth = Math.max(
      columnWidths.reduce((a, b) => a + b, 0),
      maxX + 50,
    );
    const totalHeight = Math.max(
      rowHeights.reduce((a, b) => a + b, 0),
      maxY + 50,
    );

    let shapesSvg = "";
    for (const { shape, left, top, right, bottom } of processedShapes) {
      shapesSvg += this.shapeToSvgSimple(shape, left, top, right, bottom);
    }

    return `<svg xmlns="http://www.w3.org/2000/svg" width="${totalWidth}" height="${totalHeight}" style="position:absolute;top:0;left:0;z-index:10;">
${shapesSvg}
</svg>`;
  }

  private shapeToSvgSimple(
    shape: ShapeInfo,
    x: number,
    y: number,
    right: number,
    bottom: number,
  ): string {
    let classCounter = 0;
    const getClassName = () =>
      `p${shape.name.replace(/\s/g, "_")}_${classCounter++}`;

    const width = right - x;
    const height = bottom - y;

    const schemeColorMap: Record<string, string> = {
      accent1: "#4472c4",
      accent2: "#ed7d31",
      accent3: "#a5a5a5",
      accent4: "#ffc000",
      accent5: "#5b9bd5",
      accent6: "#70ad47",
    };

    let fillColor = "white";
    let hasFill = false;
    let gradientDefs = "";
    if (shape.fill && shape.fill.type !== "none") {
      hasFill = true;
      if (shape.fill.type === "solid" && shape.fill.color) {
        if (shape.fill.isSchemeColor && schemeColorMap[shape.fill.color]) {
          fillColor = schemeColorMap[shape.fill.color];
        } else if (shape.fill.color.startsWith("#")) {
          fillColor = shape.fill.color.toUpperCase();
        } else if (schemeColorMap[shape.fill.color]) {
          fillColor = schemeColorMap[shape.fill.color];
        } else {
          fillColor = shape.fill.color.toUpperCase();
        }
      } else if (shape.fill.type === "gradient" && shape.fill.gradientStops) {
        const gradientId = `grad_${shape.name.replace(/\s/g, "_")}`;
        const angle = shape.fill.gradientAngle
          ? shape.fill.gradientAngle / 10000
          : 0;
        let gradientStopsXml = "";
        for (const stop of shape.fill.gradientStops) {
          let stopColor = stop.color;
          if (stop.isSchemeColor && schemeColorMap[stop.color]) {
            stopColor = schemeColorMap[stop.color];
          } else if (!stop.color.startsWith("#")) {
            stopColor = schemeColorMap[stop.color] || stop.color;
          }
          gradientStopsXml += `<stop offset="${stop.position / 1000}" stop-color="${stopColor}" stop-opacity="1" />`;
        }
        gradientDefs = `
  <linearGradient id="${gradientId}" x1="0%" x2="0%" y1="0%" y2="100%" gradientTransform="rotate(${angle - 90}, 0.5, 0.5)">
   ${gradientStopsXml}
  </linearGradient>`;
        fillColor = `url(#${gradientId})`;
      }
    } else if (
      shape.fillColor &&
      shape.fillColor !== "#ffffff" &&
      shape.fillColor !== "white"
    ) {
      hasFill = true;
      fillColor = shape.fillColor;
    }

    const strokeColor = shape.lineColor || "#000000";

    const type = shape.type.toLowerCase();
    const isLine = type === "line" || type.includes("connector");
    const isBentConnector =
      type.includes("bent") ||
      type.includes("elbow") ||
      type.includes("curved");

    const svgWidth = width.toFixed(2) + "pt";
    const svgHeight = height.toFixed(2) + "pt";

    let fillPath = "";
    let strokePath = "";
    let transform = "";

    const px = 0.5;
    const py = 0.5;
    const pw = width - 0.5;
    const ph = height - 0.5;

    const rotationRad = ((shape.rotation || 0) * Math.PI) / 180;
    const cosR = Math.cos(rotationRad);
    const sinR = Math.sin(rotationRad);

    if (shape.rotation && shape.rotation !== 0) {
      const cx = width / 2;
      const cy = height / 2;
      transform = `matrix(${cosR},${sinR},${-sinR},${cosR},${cy * sinR + cx - cx * cosR},${cy - cx * sinR - cy * cosR})`;
    } else if (shape.flipV) {
      transform = `matrix(1,0,0,-1,0,${height})`;
    } else if (shape.flipH) {
      transform = `matrix(-1,0,0,1,${width},0)`;
    }

    if (isLine) {
      const className = getClassName();
      strokePath = `<path d="M${x},${y} L${right},${bottom} " class="${className}" fill="none" transform="matrix(1,0,0,1,-0.000007629,-0.000007629)" />`;

      if (shape.hasArrowEnd) {
        const dx = right - x;
        const dy = bottom - y;
        const angle = Math.atan2(dy, dx);
        const arrowSize = 8;
        const arrowAngle = Math.PI / 6;

        const ax1 = right - arrowSize * Math.cos(angle - arrowAngle);
        const ay1 = bottom - arrowSize * Math.sin(angle - arrowAngle);
        const ax2 = right - arrowSize * Math.cos(angle + arrowAngle);
        const ay2 = bottom - arrowSize * Math.sin(angle + arrowAngle);

        const arrowPath = `M${right},${bottom} L${ax1},${ay1} L${ax2},${ay2} Z`;
        fillPath = `<path d="${arrowPath}" fill="${strokeColor}" transform="matrix(1,0,0,1,-0.000007629,-0.000007629)" />`;
      }
    } else if (type === "ellipse" || type === "oval") {
      const cx = px + pw / 2;
      const cy = py + ph / 2;
      const rx = pw / 2;
      const ry = ph / 2;
      const ellipsePath = `M${cx},${cy - ry} C${cx + rx},${cy - ry} ${cx + rx},${cy + ry} ${cx},${cy + ry} C${cx - rx},${cy + ry} ${cx - rx},${cy - ry} ${cx},${cy - ry} Z`;
      fillPath = hasFill
        ? `<path d="${ellipsePath}" fill="${fillColor}"${transform ? ` transform="${transform}"` : ""} />`
        : "";
      const className = getClassName();
      strokePath = `<path d="${ellipsePath}" class="${className}" fill="none"${transform ? ` transform="${transform}"` : ""} />`;
    } else if (type === "triangle" || type === "rightTriangle") {
      let pathD: string;
      if (type === "rightTriangle") {
        pathD = `M${px},${py + ph} L${px},${py} L${px + pw},${py + ph} Z`;
      } else {
        pathD = `M${px + pw / 2},${py} L${px},${py + ph} L${px + pw},${py + ph} Z`;
      }
      fillPath = hasFill
        ? `<path d="${pathD} " fill="${fillColor}"${transform ? ` transform="${transform}"` : ""} />`
        : "";
      const className = getClassName();
      strokePath = `<path d="${pathD} " class="${className}" fill="none"${transform ? ` transform="${transform}"` : ""} />`;
    } else if (type === "diamond") {
      const pathD = `M${px + pw / 2},${py} L${px + pw},${py + ph / 2} L${px + pw / 2},${py + ph} L${px},${py + ph / 2} Z`;
      fillPath = hasFill
        ? `<path d="${pathD} " fill="${fillColor}"${transform ? ` transform="${transform}"` : ""} />`
        : "";
      const className = getClassName();
      strokePath = `<path d="${pathD} " class="${className}" fill="none"${transform ? ` transform="${transform}"` : ""} />`;
    } else {
      const pathD = `M${px},${py} L${px + pw},${py} L${px + pw},${py + ph} L${px},${py + ph} Z`;
      fillPath = hasFill
        ? `<path d="${pathD} " fill="${fillColor}"${transform ? ` transform="${transform}"` : ""} />`
        : "";
      const className = getClassName();
      strokePath = `<path d="${pathD} " class="${className}" fill="none"${transform ? ` transform="${transform}"` : ""} />`;
    }

    const strokeClassName = strokePath.match(/class="([^"]+)"/)?.[1] || "";
    const styleBlock = strokeClassName
      ? `
  <defs>
   <style type="text/css">
    <![CDATA[
    
    
.${strokeClassName}
{
stroke:${strokeColor};
stroke-width:1px;
stroke-linecap:butt;
stroke-linejoin:miter;
}

    ]]></style>
  </defs>`
      : "";

    return `  <g id="SFixTitle" />
  <g id="SContent">
   <g transform="scale(1.33333)">
    <rect width="100%" height="100%" fill="white" />
    <g>
     <g>
      <g>
       ${fillPath}
       <g />
       ${strokePath}
      </g>
     </g>
    </g>
   </g>
  </g>
  <defs>${gradientDefs}
  </defs>
  <defs>${styleBlock}
</defs>
`;
  }

  private escapeHtml(str: string): string {
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }
}
