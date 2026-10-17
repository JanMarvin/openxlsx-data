library(openxlsx2)
out <- if (length(commandArgs(TRUE))) commandArgs(TRUE)[1] else "."

wb <- wb_workbook()

# 1 autofilter conditions of a table: values, custom, dynamic, color and icon filters
set.seed(1)
tf <- data.frame(
  txt = rep(c("alpha", "beta", "gamma", "delta"), 5),
  num = seq(5, 100, by = 5),
  wild = rep(c("apple pie", "banana", "cherry", "apricot"), 5),
  dyn = round(rnorm(20, 50, 10), 1),
  col = seq_len(20),
  ico = seq_len(20) * 5,
  flag = rep(c(TRUE, FALSE), 10)
)
wb$add_worksheet("tfilters")$add_data_table(x = tf, table_name = "tab_filters")
wb$add_dxfs_style("red", bg_fill = wb_color("red"))
wb$add_fill(dims = "E2:E6", color = wb_color("red"))
wb$add_conditional_formatting(dims = "F2:F21", type = "iconSet", rule = c(0, 33, 66),
                              params = list(iconSet = "3TrafficLights1"))
af <- paste0(
  '<autoFilter ref="A1:G21">',
  '<filterColumn colId="0"><filters><filter val="alpha"/><filter val="gamma"/></filters></filterColumn>',
  '<filterColumn colId="1"><customFilters and="1"><customFilter operator="greaterThan" val="10"/>',
  '<customFilter operator="lessThanOrEqual" val="80"/></customFilters></filterColumn>',
  '<filterColumn colId="2"><customFilters><customFilter operator="equal" val="a*"/>',
  '<customFilter operator="equal" val="*an*"/></customFilters></filterColumn>',
  '<filterColumn colId="3"><dynamicFilter type="aboveAverage" val="50"/></filterColumn>',
  '<filterColumn colId="4"><colorFilter dxfId="0" cellColor="1"/></filterColumn>',
  '<filterColumn colId="5"><iconFilter iconSet="3TrafficLights1" iconId="0"/></filterColumn>',
  '</autoFilter>'
)
wb$tables$tab_xml[1] <- sub('<autoFilter ref="A1:G21"/>', af, wb$tables$tab_xml[1], fixed = TRUE)

# 2 autofilter conditions of a worksheet: more variants, dates, blanks, top10 and bottom10, percent
wd <- data.frame(
  a = c(1:18, NA, NA),
  b = rep(c("x", "y", "z", ""), 5),
  c = as.Date("2020-01-15") + seq(0, by = 40, length.out = 20),
  d = (1:20)^2,
  e = rep(c("Mazda", "Fiat", "Merc", "Audi"), 5),
  f = 20:1
)
wb$add_worksheet("wfilters")$add_data(x = wd)
wb$add_numfmt(dims = "C2:C21", numfmt = "yyyy-mm-dd")
wb$worksheets[[2]]$autoFilter <- paste0(
  '<autoFilter ref="A1:F21">',
  '<filterColumn colId="0"><filters blank="1"><filter val="3"/><filter val="7"/></filters></filterColumn>',
  '<filterColumn colId="1"><customFilters><customFilter operator="notEqual" val="x"/></customFilters></filterColumn>',
  '<filterColumn colId="2"><filters><dateGroupItem year="2020" month="1" dateTimeGrouping="month"/>',
  '<dateGroupItem year="2020" month="3" day="5" dateTimeGrouping="day"/></filters></filterColumn>',
  '<filterColumn colId="3"><top10 top="0" percent="1" val="20" filterVal="16"/></filterColumn>',
  '<filterColumn colId="4"><customFilters and="1"><customFilter operator="notEqual" val="Fiat"/>',
  '<customFilter operator="notEqual" val="Audi"/></customFilters></filterColumn>',
  '<filterColumn colId="5"><dynamicFilter type="belowAverage" val="10.5"/></filterColumn>',
  '</autoFilter>'
)

# 3 strings: newlines, escapes, non ASCII, surrogate pairs, in cells, formulas and a sheet name
st <- data.frame(
  text = c("line1\nline2", "tab\there", "a & b < c > d", "quote \" and '", "\u00e4\u00f6\u00fc\u00df",
           "\u65e5\u672c\u8a9e", "\U0001F600 smile", "\U0001D11E clef", "trailing space ", "  leading space"),
  stringsAsFactors = FALSE
)
wb$add_worksheet("strings")$add_data(x = st)
wb$add_formula(dims = "B2", x = '"first"&CHAR(10)&"second"')
wb$add_formula(dims = "B3", x = 'A2&" - "&A3')
wb$add_formula(dims = "B4", x = 'LEN(A5)')
wb$add_formula(dims = "B5", x = 'CODE(A8)')
wb$add_formula(dims = "B12", x = 'IF(1=1,"multi\nline","other")')
wb$add_worksheet("\u00dcml\u00e4ut \U0001F600")$add_data(x = 1:3)
wb$add_named_region(sheet = "strings", dims = "A2:A4", name = "unicode_\u00e4")

# 4 formulas and names with reference errors
wb$add_worksheet("errors")$add_data(x = 1:10)
wb$add_formula(dims = "B1", x = "#REF!")
wb$add_formula(dims = "B2", x = "SUM(#REF!)")
wb$add_formula(dims = "B3", x = "#REF!+A3")
wb$add_formula(dims = "B4", x = "SUM(#REF!:A4)")
wb$add_formula(dims = "B5", x = "SUM(A5:#REF!)")
wb$add_formula(dims = "B6", x = "SUM(#REF!,A6)")
wb$add_formula(dims = "B7", x = "SUM(errors!#REF!)")
wb$add_formula(dims = "B8", x = "#DIV/0!")
wb$add_formula(dims = "B9", x = "#N/A")
wb$add_formula(dims = "B10", x = "IFERROR(#NAME?,0)")
wb$add_named_region(sheet = "errors", dims = "A1:A3", name = "gone")
wb$workbook$definedNames <- gsub("'errors'!\\$A\\$1:\\$A\\$3", "#REF!", wb$workbook$definedNames)

# 5 external workbook reference, including a blank cell in the cached values
wb$add_worksheet("external")$add_data(x = 1:5)
wb$add_formula(dims = "B1", x = "'[1]Sheet1'!$A$1")
wb$add_formula(dims = "B2", x = "'[1]Sheet1'!$A$2")
wb$add_formula(dims = "B3", x = "'[1]Sheet1'!$A$3")
wb$add_formula(dims = "B4", x = "'[1]Sheet1'!$B$1:$B$3")
wb$add_formula(dims = "B5", x = "SUM('[1]Sheet1'!$A$1:$A$3)")
wb$externalLinks <- list(paste0(
  '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
  '<externalBook r:id="rId1"><sheetNames><sheetName val="Sheet1"/></sheetNames>',
  '<sheetDataSet><sheetData sheetId="0">',
  '<row r="1"><cell r="A1"><v>1</v></cell><cell r="B1"/></row>',
  '<row r="2"><cell r="A2" t="str"><v>text</v></cell></row>',
  '<row r="3"><cell r="A3" t="b"><v>1</v></cell><cell r="B3" t="e"><v>#N/A</v></cell></row>',
  '</sheetData></sheetDataSet></externalBook></externalLink>'
))
wb$externalLinksRels <- list(paste0(
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" ',
  'Target="ext_cov4.xlsx" TargetMode="External"/></Relationships>'
))

# 6 custom filters on booleans and mixed values
bf <- data.frame(flag = rep(c(TRUE, FALSE, TRUE, TRUE), 5), other = rep(c(FALSE, TRUE), 10), n = 1:20)
wb$add_worksheet("bool_filter")$add_data(x = bf)
wb$worksheets[[7]]$autoFilter <- paste0(
  '<autoFilter ref="A1:C21">',
  '<filterColumn colId="0"><customFilters><customFilter operator="equal" val="TRUE"/></customFilters></filterColumn>',
  '<filterColumn colId="1"><filters><filter val="1"/></filters></filterColumn>',
  '</autoFilter>'
)

wb$workbook$externalReferences <- '<externalReferences><externalReference r:id="rId99"/></externalReferences>'
wb$workbook.xml.rels <- c(wb$workbook.xml.rels, paste0(
  '<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" ',
  'Target="externalLinks/externalLink1.xml"/>'
))
wb$Content_Types <- c(wb$Content_Types, paste0(
  '<Override PartName="/xl/externalLinks/externalLink1.xml" ',
  'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>'
))

wb$save(file.path(out, "xls_cov4_misc.xlsx"))
