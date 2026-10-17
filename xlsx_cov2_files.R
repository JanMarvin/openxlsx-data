library(openxlsx2)
out <- if (length(commandArgs(TRUE))) commandArgs(TRUE)[1] else "."

# 1 pivot tables: functions, show data as, sort, hidden items, totals, layouts
wb <- wb_workbook()$add_worksheet("src")$add_data(x = mtcars)
df <- wb_data(wb, sheet = "src")

wb$add_worksheet("functions")
wb$add_pivot_table(x = df, sheet = "functions", dims = "A3", rows = "cyl", cols = "am", data = c("mpg", "hp", "wt"), fun = c("average", "max", "min"))
wb$add_worksheet("functions2")
wb$add_pivot_table(x = df, sheet = "functions2", dims = "A3", rows = "cyl", data = c("mpg", "hp", "wt", "qsec"), fun = c("product", "stdDev", "var", "count"))
wb$add_worksheet("functions3")
wb$add_pivot_table(x = df, sheet = "functions3", dims = "A3", rows = "gear", data = c("mpg", "hp", "wt"), fun = c("stdDevp", "varp", "countNums"))

shows <- c("normal", "runTotal", "percentOfRow", "percentOfCol", "percentOfTotal", "index")
for (i in seq_along(shows)) {
  sh <- paste0("show_", shows[i])
  wb$add_worksheet(sh)
  wb$add_pivot_table(x = df, sheet = sh, dims = "A3", rows = "cyl", cols = "am", data = "mpg", fun = "sum",
                     params = list(show_data_as = shows[i]))
}

wb$add_worksheet("sorted")
wb$add_pivot_table(x = df, sheet = "sorted", dims = "A3", rows = c("cyl", "gear"), cols = "am", data = "mpg",
                   params = list(sort_row = "descending", sort_col = "ascending"))
wb$add_worksheet("sorted2")
wb$add_pivot_table(x = df, sheet = "sorted2", dims = "A3", rows = c("gear", "cyl"), cols = c("am", "vs"), data = c("mpg", "hp"),
                   params = list(sort_row = "ascending", sort_col = "descending"))

wb$add_worksheet("totals")
wb$add_pivot_table(x = df, sheet = "totals", dims = "A3", rows = c("cyl", "gear"), cols = "am", data = "mpg",
                   params = list(row_grand_totals = FALSE, col_grand_totals = FALSE, default_subtotal = FALSE))
wb$add_worksheet("totals2")
wb$add_pivot_table(x = df, sheet = "totals2", dims = "A3", rows = c("cyl", "gear"), cols = "am", data = "mpg",
                   params = list(row_grand_totals = FALSE, subtotal_top = FALSE))
wb$add_worksheet("compact")
wb$add_pivot_table(x = df, sheet = "compact", dims = "A3", rows = c("cyl", "gear"), data = "mpg",
                   params = list(compact = FALSE, outline = FALSE))
wb$add_worksheet("outline")
wb$add_pivot_table(x = df, sheet = "outline", dims = "A3", rows = c("cyl", "gear"), data = "mpg",
                   params = list(compact = FALSE, outline = TRUE, compact_data = FALSE))

wb$add_worksheet("choose")
wb$add_pivot_table(x = df, sheet = "choose", dims = "A3", rows = "cyl", cols = "am", data = "mpg", filter = "gear",
                   params = list(choose = c(cyl = 'x > "4"')))
wb$add_worksheet("filters")
wb$add_pivot_table(x = df, sheet = "filters", dims = "A3", rows = "cyl", data = "mpg", filter = c("gear", "carb", "am"))
wb$add_worksheet("cols_only")
wb$add_pivot_table(x = df, sheet = "cols_only", dims = "A3", cols = c("cyl", "gear"), data = "mpg")
wb$add_worksheet("rows_multi_data")
wb$add_pivot_table(x = df, sheet = "rows_multi_data", dims = "A3", rows = c("cyl", "am", "vs"), data = c("mpg", "hp", "wt"))
wb$add_worksheet("calc_field")
wb$add_pivot_table(x = df, sheet = "calc_field", dims = "A3", rows = "cyl", data = c("am", "=disp/cyl" = "New"))
wb$add_worksheet("numfmt")
wb$add_pivot_table(x = df, sheet = "numfmt", dims = "A3", rows = "cyl", cols = "am", data = "mpg",
                   params = list(numfmts = c(formatCode = "0.0%")))
wb$add_worksheet("styled")
wb$add_pivot_table(x = df, sheet = "styled", dims = "A3", rows = "cyl", cols = "am", data = "mpg",
                   params = list(table_style = "TableStyleMedium23", apply_number_formats = TRUE, apply_border_formats = TRUE,
                                 apply_font_formats = TRUE, apply_pattern_formats = TRUE, apply_alignment_formats = TRUE))
wb$add_worksheet("nostyle")
wb$add_pivot_table(x = df, sheet = "nostyle", dims = "A3", rows = "cyl", cols = "am", data = "mpg", params = list(no_style = TRUE))

wb$add_worksheet("text_src")$add_data(x = data.frame(
  region = rep(c("north", "south", "east", "west"), each = 6),
  product = rep(c("a", "b", "c"), 8),
  day = rep(as.Date("2024-01-01") + 0:5, 4),
  amount = (1:24) * 1.5,
  stringsAsFactors = FALSE))
tdf <- wb_data(wb, sheet = "text_src")
wb$add_worksheet("text_pivot")
wb$add_pivot_table(x = tdf, sheet = "text_pivot", dims = "A3", rows = c("region", "product"), cols = "day", data = "amount")
wb$add_worksheet("text_pivot2")
wb$add_pivot_table(x = tdf, sheet = "text_pivot2", dims = "A3", rows = "day", filter = "region", data = "amount", fun = "count")
wb$save(file.path(out, "xls_cov2_pivots.xlsx"))

# 2 values: every error, bool, built-in number formats, unusual numbers
wb <- wb_workbook()$add_worksheet("errors")
err <- c("NA()", "1/0", "\"a\"+1", "SQRT(-1)", "NOSUCHFUNCTION(1)", "A1 B2", "INDEX(A1:A2,5)", "LOG(0)", "ISERROR(1/0)", "IFERROR(1/0,\"ok\")", "TRUE", "FALSE", "AND(TRUE,FALSE)", "NOT(FALSE)", "ERROR.TYPE(NA())", "IF(1>2,1,\"s\")&\"x\"", "\"\"", "T(1)", "REPT(\"x\",300)", "REPT(\"y\",5000)")
for (i in seq_along(err)) wb$add_formula(sheet = "errors", x = err[i], dims = wb_dims(rows = i, cols = 1))
wb$add_data(sheet = "errors", x = c(1, NA, 3), dims = "C1", col_names = FALSE)

wb$add_worksheet("numfmts")
ids <- c(0:22, 37:40, 45:49)
for (i in seq_along(ids)) {
  wb$add_data(sheet = "numfmts", x = c(1234.5678, 45000.75, -0.5)[(i %% 3) + 1], dims = wb_dims(rows = i, cols = 1), col_names = FALSE)
  wb$add_numfmt(sheet = "numfmts", dims = wb_dims(rows = i, cols = 1), numfmt = ids[i])
  wb$add_data(sheet = "numfmts", x = ids[i], dims = wb_dims(rows = i, cols = 2), col_names = FALSE)
}
wb$add_worksheet("numbers")
wb$add_data(sheet = "numbers", x = c(0, 1, -1, 100, 536870911, 536870912, -536870912, -536870913, 0.1, 1/3, 1e15, 1e-15, 123456.789, 2^31, 2^53, -2^31, 255, 65535, 65536, 0.25, 0.125), col_names = FALSE)
wb$add_data(sheet = "numbers", x = c(TRUE, FALSE, NA, TRUE), dims = "B1", col_names = FALSE)
wb$add_data(sheet = "numbers", x = as.Date("1900-03-01") + c(0, 1, 100, 40000, 50000, -1000), dims = "C1", col_names = FALSE)
wb$add_data(sheet = "numbers", x = as.POSIXct("1999-12-31 23:59:59", tz = "UTC") + c(0, 1, 86400, 1e6), dims = "D1", col_names = FALSE)
wb$add_data(sheet = "numbers", x = c("a", "ä", "日本語", "😀", "x\ny", "  lead", "trail  ", "tab\t", strrep("z", 1000), "'quote", "=notformula", "<tag>&amp;"), dims = "E1", col_names = FALSE)
wb$add_worksheet("colors")
cols <- c("#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF", "#800000", "#008000", "#000080", "#808000", "#800080", "#008080", "#C0C0C0", "#808080", "#9999FF", "#993366", "#FFFFCC", "#CCFFFF", "#660066", "#FF8080", "#0066CC", "#CCCCFF", "#123456", "#ABCDEF", "#FEDCBA", "#102030")
wb$add_data(sheet = "colors", x = seq_along(cols), dims = "A1", col_names = FALSE)
for (i in seq_along(cols)) {
  wb$add_fill(sheet = "colors", dims = wb_dims(rows = i, cols = 1), color = wb_color(cols[i]))
  wb$add_font(sheet = "colors", dims = wb_dims(rows = i, cols = 1), color = wb_color(rev(cols)[i]))
  wb$add_border(sheet = "colors", dims = wb_dims(rows = i, cols = 2), top_border = "thin", top_color = wb_color(cols[i]), bottom_border = "medium", bottom_color = wb_color(rev(cols)[i]))
}
wb$add_worksheet("fonts")
fn <- c("Arial", "Times New Roman", "Courier New", "Calibri", "Verdana", "Tahoma", "Georgia", "Comic Sans MS", "Impact", "Trebuchet MS")
for (i in seq_along(fn)) {
  wb$add_data(sheet = "fonts", x = fn[i], dims = wb_dims(rows = i, cols = 1), col_names = FALSE)
  wb$add_font(sheet = "fonts", dims = wb_dims(rows = i, cols = 1), name = fn[i], size = 8 + i, bold = i %% 2 == 0, italic = i %% 3 == 0,
              underline = c("single", "double", "singleAccounting", "doubleAccounting")[i %% 4 + 1], strike = i %% 5 == 0)
}
wb$save(file.path(out, "xls_cov2_values.xlsx"))

# 3 defined names
wb <- wb_workbook()$add_worksheet("a")$add_data(x = mtcars, row_names = TRUE)
wb$add_worksheet("b")$add_data(x = iris)
wb$add_worksheet("c with space")$add_data(x = 1:10)
wb$workbook$definedNames <- c(
  '<definedName name="plain">a!$A$1:$B$2</definedName>',
  '<definedName name="other_sheet">b!$A$1:$B$2</definedName>',
  '<definedName name="space_sheet">\'c with space\'!$A$1:$A$10</definedName>',
  '<definedName name="local_a" localSheetId="0">a!$A$1</definedName>',
  '<definedName name="local_a" localSheetId="1">b!$A$1</definedName>',
  '<definedName name="local_hidden" localSheetId="2" hidden="1">\'c with space\'!$A$1</definedName>',
  '<definedName name="refs_name">plain</definedName>',
  '<definedName name="sum_name">SUM(plain)+SUM(other_sheet)</definedName>',
  '<definedName name="rel_r1c1">OFFSET(a!$A$1,ROW()-1,0)</definedName>',
  '<definedName name="str_const">"a long text constant"</definedName>',
  '<definedName name="num_const">42</definedName>',
  '<definedName name="bool_const">TRUE</definedName>',
  '<definedName name="err_const">#N/A</definedName>',
  '<definedName name="arr_const">{1,2,3}</definedName>',
  '<definedName name="arr_str">{"a","b";"c","d"}</definedName>',
  '<definedName name="multi_area">(a!$A$1:$A$2,b!$A$1:$A$2)</definedName>',
  '<definedName name="three_d">a:c!$A$1</definedName>',
  '<definedName name="whole_col">a!$A:$A</definedName>',
  '<definedName name="whole_row">a!$1:$1</definedName>',
  '<definedName name="bad_ref">#REF!</definedName>',
  '<definedName name="ext_ref">[xls_cov2_values.xlsx]errors!$A$1</definedName>',
  '<definedName name="described" comment="a comment" description="a description">a!$A$1</definedName>',
  '<definedName name="_xlnm.Print_Area" localSheetId="0">a!$A$1:$F$20</definedName>',
  '<definedName name="_xlnm.Print_Area" localSheetId="1">b!$A$1:$C$10,b!$E$1:$F$10</definedName>',
  '<definedName name="_xlnm.Print_Titles" localSheetId="0">a!$A:$B,a!$1:$2</definedName>',
  '<definedName name="_xlnm.Print_Titles" localSheetId="1">b!$1:$1</definedName>',
  '<definedName name="_xlnm.Print_Titles" localSheetId="2">\'c with space\'!$A:$A</definedName>',
  '<definedName name="_xlnm._FilterDatabase" localSheetId="1" hidden="1">b!$A$1:$E$151</definedName>',
  '<definedName name="_xlnm.Criteria" localSheetId="0">a!$N$1:$O$2</definedName>',
  '<definedName name="_xlnm.Extract" localSheetId="0">a!$N$5:$P$5</definedName>',
  '<definedName name="_xlnm.Consolidate_Area" localSheetId="0">a!$R$1:$S$5</definedName>',
  '<definedName name="_xlnm.Sheet_Title" localSheetId="0">"title"</definedName>'
)
wb$add_filter(sheet = "b", rows = 1, cols = 1:5)
wb$save(file.path(out, "xls_cov2_names.xlsx"))

# 4 shapes: one workbook sheet with every preset geometry the importer knows, plus text boxes, lines, rotation, groups
presets <- c("rect", "roundRect", "ellipse", "diamond", "triangle", "rtTriangle", "parallelogram", "trapezoid", "hexagon", "octagon", "plus", "star5", "rightArrow", "homePlate", "cube", "line", "plaque", "can", "donut", "pentagon", "chevron", "star8", "star16", "star24", "star32", "wedgeRectCallout", "wedgeRoundRectCallout", "leftArrow", "upArrow", "downArrow", "leftRightArrow", "upDownArrow", "bentArrow", "curvedRightArrow", "stripedRightArrow", "notchedRightArrow", "bevel", "foldedCorner", "smileyFace", "noSmoking", "flowChartProcess", "flowChartTerminator", "flowChartDecision", "flowChartInputOutput", "sun", "moon", "bracketPair", "bracePair", "star4", "doubleWave", "blockArc", "heart", "lightningBolt", "star6")
wb <- wb_workbook()$add_worksheet("shapes")
extra <- c("heptagon", "decagon", "teardrop", "pie", "chord", "arc", "cloud", "cloudCallout", "ribbon", "wave", "gear6", "funnel", "mathPlus", "frame", "uturnArrow", "quadArrow", "flowChartDocument", "actionButtonHome", "irregularSeal1", "verticalScroll")
fills <- c("FF0000", "00B050", "0070C0", "FFC000", "7030A0", "A5A5A5")
aligns <- c("left", "center", "right")
all_shapes <- c(presets, extra)
for (i in seq_along(all_shapes)) {
  col <- (i - 1L) %% 6L
  row <- (i - 1L) %/% 6L
  xml <- create_shape(
    shape = all_shapes[i],
    name = paste("shape", i),
    text = all_shapes[i],
    fill_color = wb_color(fills[col + 1L]),
    fill_transparency = if (i %% 5L == 0L) 50 else 0,
    text_color = wb_color("FFFFFF"),
    line_color = wb_color("000000"),
    text_align = aligns[i %% 3L + 1L],
    rotation = if (i %% 7L == 0L) 30 else 0,
    id = i + 1L
  )
  wb$add_drawing(sheet = "shapes", dims = wb_dims(rows = (2L + row * 4L):(4L + row * 4L), cols = (2L + col * 3L):(3L + col * 3L)), xml = xml)
}
wb$add_worksheet("textboxes")
tb <- c("plain text box", "second line", "third")
for (i in 1:6) {
  xml <- paste0(
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">',
    '<xdr:absoluteAnchor><xdr:pos x="400000" y="', sprintf("%d", 300000L + (i - 1L) * 1300000L), '"/><xdr:ext cx="3000000" cy="1000000"/>',
    '<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="', i + 1, '" name="TextBox ', i, '"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr>',
    '<xdr:spPr><a:xfrm><a:off x="400000" y="', sprintf("%d", 300000L + (i - 1L) * 1300000L), '"/><a:ext cx="3000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>',
    if (i %% 2 == 0) '<a:solidFill><a:srgbClr val="FFFFCC"/></a:solidFill>' else '<a:noFill/>',
    '<a:ln w="9525"><a:solidFill><a:srgbClr val="', if (i %% 2 == 0) 'FF0000' else '0070C0', '"/></a:solidFill></a:ln></xdr:spPr>',
    '<xdr:txBody><a:bodyPr vertOverflow="clip" wrap="', if (i %% 3 == 0) 'none' else 'square', '" rtlCol="0" anchor="', c("t", "ctr", "b")[i %% 3 + 1], '"/><a:lstStyle/>',
    '<a:p><a:pPr algn="', c("l", "ctr", "r")[i %% 3 + 1], '"/><a:r><a:rPr lang="en-US" sz="', 800 + i * 200, '" b="', i %% 2, '" i="', (i + 1) %% 2, '" u="', if (i %% 2) 'sng' else 'none', '"><a:solidFill><a:srgbClr val="', c("000000", "FF0000", "00B050")[i %% 3 + 1], '"/></a:solidFill><a:latin typeface="', c("Arial", "Times New Roman", "Courier New")[i %% 3 + 1], '"/></a:rPr><a:t>', tb[i %% 3 + 1], '</a:t></a:r></a:p>',
    '<a:p><a:r><a:rPr lang="en-US" sz="1100"/><a:t>second paragraph ', i, ' ä€</a:t></a:r></a:p></xdr:txBody>',
    '</xdr:sp><xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>')
  wb$add_drawing(sheet = "textboxes", xml = xml)
}
wb$add_worksheet("connectors")
for (i in 1:6) {
  y <- sprintf("%d", 300000L + (i - 1L) * 700000L)
  xml <- paste0(
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">',
    '<xdr:absoluteAnchor><xdr:pos x="400000" y="', y, '"/><xdr:ext cx="3000000" cy="400000"/>',
    '<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="', 100L + i, '" name="Group ', i, '"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>',
    '<xdr:grpSpPr><a:xfrm><a:off x="400000" y="', y, '"/><a:ext cx="3000000" cy="400000"/><a:chOff x="400000" y="', y, '"/><a:chExt cx="3000000" cy="400000"/></a:xfrm></xdr:grpSpPr>',
    '<xdr:cxnSp macro=""><xdr:nvCxnSpPr><xdr:cNvPr id="', i + 1L, '" name="Connector ', i, '"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr>',
    '<xdr:spPr><a:xfrm', if (i %% 2 == 0) ' flipV="1"' else '', '><a:off x="400000" y="', y, '"/><a:ext cx="3000000" cy="400000"/></a:xfrm><a:prstGeom prst="', c("line", "straightConnector1", "bentConnector3", "curvedConnector3", "line", "straightConnector1")[i], '"><a:avLst/></a:prstGeom>',
    '<a:ln w="', 9525L * i, '"><a:solidFill><a:srgbClr val="', fills[i], '"/></a:solidFill>', if (i > 2) '<a:prstDash val="sysDot"/>' else '', if (i %% 2 == 0) '<a:tailEnd type="triangle"/>' else '<a:headEnd type="oval"/>', '</a:ln></xdr:spPr></xdr:cxnSp>',
    '</xdr:grpSp><xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>')
  wb$add_drawing(sheet = "connectors", xml = xml)
}
wb$save(file.path(out, "xls_cov2_shapes.xlsx"))
