library(openxlsx2)
out <- if (length(commandArgs(TRUE))) commandArgs(TRUE)[1] else "."

# 1 view: panes, zoom, tab color, protection, grid/headers, hidden rows/cols, outlines
wb <- wb_workbook()$add_worksheet("freeze", tab_color = wb_color("red"), zoom = 85)$
  add_data(x = head(mtcars, 30), row_names = TRUE)$
  freeze_pane(first_active_row = 3, first_active_col = 2)
wb$add_worksheet("freeze_row", grid_lines = FALSE, tab_color = wb_color("#00AA55"))$
  add_data(x = head(iris, 20))$freeze_pane(first_row = TRUE)
wb$add_worksheet("freeze_col")$add_data(x = head(iris, 20))$freeze_pane(first_col = TRUE)
wb$add_worksheet("split", zoom = 150)$add_data(x = head(iris, 20))
wb$worksheets[[4]]$sheetViews <- '<sheetViews><sheetView workbookViewId="0" zoomScale="150"><pane xSplit="2400" ySplit="1800" topLeftCell="C4" activePane="bottomRight" state="split"/><selection pane="topRight"/><selection pane="bottomLeft"/><selection pane="bottomRight" activeCell="D5" sqref="D5"/></sheetView></sheetViews>'
wb$add_worksheet("hidden_outline")$add_data(x = head(iris, 20))$
  group_rows(rows = 3:6, collapsed = TRUE)$group_cols(cols = 2:3)$
  set_row_heights(rows = 8, heights = 40)$set_col_widths(cols = 1:5, widths = c(5, 12, 20, "auto", 30))
wb$add_worksheet("protected")$add_data(x = head(iris, 5))$
  protect_worksheet(protect = TRUE, password = "pw", properties = c("formatCells", "insertRows"))
wb$add_worksheet("very_hidden", visible = FALSE)$add_data(x = 1)
wb$set_bookview(active_tab = 1, first_sheet = 0)
wb$protect(protect = TRUE, password = "pw", lock_structure = TRUE)
wb$save(file.path(out, "xls_cov_views.xlsx"))

# 2 filters, merges, validation, conditional formats, hyperlinks, page setup
wb <- wb_workbook()
wb$add_worksheet("filter")$add_data(x = mtcars, row_names = TRUE)$
  add_filter(rows = 1, cols = 1:12)
wb$worksheets[[1]]$autoFilter <- '<autoFilter ref="A1:L33"><filterColumn colId="2"><filters><filter val="6"/><filter val="8"/></filters></filterColumn><filterColumn colId="3"><customFilters and="1"><customFilter operator="greaterThan" val="100"/><customFilter operator="lessThan" val="300"/></customFilters></filterColumn><filterColumn colId="4"><top10 top="1" val="5"/></filterColumn></autoFilter>'
wb$add_worksheet("merge_val")$add_data(x = head(iris, 10))$
  merge_cells(dims = "G1:H3")$merge_cells(dims = "G5:J5")$
  add_data_validation(dims = "A2:A10", type = "whole", operator = "between", value = c(1, 10), allow_blank = FALSE, show_input_msg = TRUE, prompt_title = "t", prompt = "p", show_error_msg = TRUE, error_title = "e", error = "bad")$
  add_data_validation(dims = "E2:E10", type = "list", value = '"setosa,versicolor,virginica"')$
  add_data_validation(dims = "B2:B10", type = "decimal", operator = "greaterThan", value = 1)$
  add_data_validation(dims = "C2:C10", type = "date", operator = "greaterThan", value = as.Date("2020-01-01"))$
  add_data_validation(dims = "D2:D10", type = "textLength", operator = "lessThan", value = 5)$
  add_data_validation(dims = "F2:F10", type = "custom", value = "F2>0")
wb$add_dxfs_style("d1", font_color = wb_color("red"), bg_fill = wb_color("yellow"))
wb$add_dxfs_style("d2", bg_fill = wb_color("orange"), border = TRUE, border_color = wb_color("blue"), border_style = "thick")
wb$add_dxfs_style("d3", text_bold = TRUE, text_italic = TRUE, text_underline = "single", num_fmt = 2, font_name = "Arial", font_size = 14)
wb$add_dxfs_style("d4", bg_fill = wb_color("pink"), format_code = "0.00%")
wb$add_worksheet("cf")$add_data(x = head(mtcars, 20), row_names = TRUE)$
  add_conditional_formatting(dims = "B2:B21", rule = ">20", style = "d1")$
  add_conditional_formatting(dims = "C2:C21", type = "colorScale", style = c("blue", "white", "red"), rule = c(0, 5, 10))$
  add_conditional_formatting(dims = "D2:D21", type = "colorScale", style = c("green", "yellow"), rule = c(50, 300))$
  add_conditional_formatting(dims = "E2:E21", type = "dataBar", style = c("green", "lightgreen"))$
  add_conditional_formatting(dims = "F2:F21", type = "iconSet", rule = c(0, 10, 20), params = list(iconSet = "3TrafficLights1", reverse = FALSE))$
  add_conditional_formatting(dims = "G2:G21", type = "iconSet", rule = c(0, 20, 40, 60, 80), params = list(iconSet = "5Arrows", reverse = TRUE, showValue = FALSE))$
  add_conditional_formatting(dims = "H2:H21", type = "topN", params = list(rank = 3), style = "d2")$
  add_conditional_formatting(dims = "I2:I21", type = "bottomN", params = list(rank = 10, percent = TRUE), style = "d4")$
  add_conditional_formatting(dims = "J2:J21", type = "duplicatedValues", style = "d4")$
  add_conditional_formatting(dims = "K2:K21", type = "uniqueValues", style = "d2")$
  add_conditional_formatting(dims = "A2:A21", type = "containsText", rule = "Mazda", style = "d1")$
  add_conditional_formatting(dims = "A2:A21", type = "notContainsText", rule = "Fiat", style = "d3")$
  add_conditional_formatting(dims = "A2:A21", type = "beginsWith", rule = "Merc", style = "d3")$
  add_conditional_formatting(dims = "A2:A21", type = "endsWith", rule = "4", style = "d2")$
  add_conditional_formatting(dims = "L2:L21", type = "expression", rule = "L2>$B2", style = "d3")$
  add_conditional_formatting(dims = "B2:B21", type = "between", rule = c(1, 3), style = "d3")$
  add_conditional_formatting(dims = "M2:M21", type = "containsBlanks", style = "d1")$
  add_conditional_formatting(dims = "N2:N21", type = "containsErrors", style = "d1")
wb$add_worksheet("links")$add_data(x = c("a", "b", "c", "d"))$
  add_hyperlink(dims = "A1", target = "https://www.r-project.org", tooltip = "R project")$
  add_hyperlink(dims = "A2", target = "mailto:someone@example.com", tooltip = "mail", is_external = TRUE)$
  add_hyperlink(dims = "A3", target = "'filter'!B2", is_external = FALSE, tooltip = "internal")$
  add_hyperlink(dims = "A4", target = "file:///C:/temp/test.xlsx", tooltip = "file")
wb$add_worksheet("page", header = c("left h", "&[Tab]", "right h"), footer = c("&[Page] of &[Pages]", "&[File]", "&[Date]"),
                 even_header = c("e1", "e2", "e3"), first_header = c("f1", "f2", "f3"),
                 orientation = "landscape", paper_size = 9)$
  add_data(x = mtcars, row_names = TRUE)$
  add_page_break(row = 20)$add_page_break(col = 5)$
  set_page_setup(scale = 80, left = 0.5, right = 0.5, top = 1, bottom = 1, header = 0.4, footer = 0.4,
                 print_title_rows = 1, print_title_cols = 1, black_and_white = TRUE, draft = TRUE,
                 horizontal_centered = TRUE, vertical_centered = TRUE, print_headings = TRUE,
                 cell_comments = "atEnd", errors = "blank", page_order = "overThenDown", first_page_number = 3, use_first_page_number = TRUE)
wb$add_worksheet("fit")$add_data(x = mtcars)$set_page_setup(fit_to_width = 1, fit_to_height = 0, orientation = "portrait", paper_size = 8)
wb$add_named_region(dims = "A1:B3", name = "print_area_like", sheet = "page")
wb$save(file.path(out, "xls_cov_features.xlsx"))

# 3 defined names, formulas, external, calc settings
wb <- wb_workbook()
wb$add_worksheet("data")$add_data(x = head(mtcars, 20), row_names = TRUE)
wb$add_worksheet("calc")$add_data(x = 1:10)
wb$add_named_region(dims = "A1:B5", name = "rng", sheet = "data")
wb$add_named_region(dims = "A1:B5", name = "local_rng", sheet = "calc", local_sheet = TRUE)
wb$add_named_region(dims = "A1:A3", name = "hidden_rng", sheet = "calc")
wb$workbook$definedNames <- c(wb$workbook$definedNames,
                              '<definedName name="const_num">3.14159</definedName>',
                              '<definedName name="const_str">"text"</definedName>',
                              '<definedName name="const_arr">{1,2,3;4,5,6}</definedName>',
                              '<definedName name="fml_name">SUM(data!$B$2:$B$21)*2</definedName>',
                              '<definedName name="ofs">OFFSET(data!$A$1,0,0,COUNTA(data!$A:$A),1)</definedName>',
                              '<definedName name="hid" hidden="1">data!$C$2</definedName>',
                              '<definedName name="cmt" comment="a comment" description="desc" function="0">data!$D$2</definedName>',
                              '<definedName name="_xlnm.Print_Area" localSheetId="0">data!$A$1:$E$10</definedName>',
                              '<definedName name="_xlnm.Print_Titles" localSheetId="0">data!$A:$A,data!$1:$1</definedName>',
                              '<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">data!$A$1:$L$21</definedName>',
                              '<definedName name="multi">(data!$A$1,data!$B$1:$B$3)</definedName>',
                              '<definedName name="three_d">data:calc!$A$1</definedName>')
f <- c(
  "SUM(data!B2:B21)", "AVERAGE(rng)", "IF(A1>5,\"hi\",\"lo\")", "VLOOKUP(\"Mazda RX4\",data!A1:L21,2,FALSE)",
  "INDEX(data!B2:D21,MATCH(\"Valiant\",data!A2:A21,0),2)", "SUMPRODUCT((data!B2:B21>6)*(data!C2:C21))",
  "TEXT(A1,\"0.00\")&\"x\"", "CONCATENATE(\"a\",\"b\")", "LEFT(\"abcdef\",3)&MID(\"abcdef\",2,2)&RIGHT(\"abcdef\",1)",
  "ROUND(PI(),3)", "DATE(2020,1,31)+EDATE(DATE(2020,1,31),1)", "NETWORKDAYS(DATE(2020,1,1),DATE(2020,2,1))",
  "IFERROR(1/0,\"err\")", "ISNUMBER(A1)", "AND(TRUE,FALSE)", "OR(TRUE,FALSE)", "NOT(TRUE)", "CHOOSE(2,\"a\",\"b\",\"c\")",
  "COUNTIF(A1:A10,\">5\")", "SUMIF(A1:A10,\">5\",A1:A10)", "SUMIFS(A1:A10,A1:A10,\">2\",A1:A10,\"<8\")",
  "COUNTIFS(A1:A10,\">2\",A1:A10,\"<8\")", "MAX(A1:A10)-MIN(A1:A10)", "LARGE(A1:A10,2)", "RANK(A2,A1:A10)",
  "STDEV(A1:A10)", "MEDIAN(A1:A10)", "MOD(A1,3)", "POWER(2,10)", "SQRT(16)", "ABS(-3)", "INT(2.7)", "LN(EXP(1))",
  "FIND(\"b\",\"abc\")", "SUBSTITUTE(\"aaa\",\"a\",\"b\")", "LEN(\"abc\")", "TRIM(\" a \")", "UPPER(\"a\")&LOWER(\"B\")&PROPER(\"cd\")",
  "ROW()+COLUMN()", "ROWS(A1:A10)+COLUMNS(A1:C1)", "OFFSET(A1,1,1)", "INDIRECT(\"A1\")", "ADDRESS(1,1)", "HYPERLINK(\"http://x.org\",\"x\")",
  "TODAY()-TODAY()", "NOW()-NOW()", "YEAR(DATE(2020,1,1))+MONTH(DATE(2020,2,1))+DAY(DATE(2020,3,1))", "WEEKDAY(DATE(2020,1,1))",
  "-A1", "A1%", "A1^2", "A1&A2", "A1=A2", "A1<>A2", "A1>=A2", "A1<=A2", "(A1+A2)*(A3-A4)/2", "SUM(A1:A3 A2:A4)", "SUM((A1:A2,A3:A4))",
  "calc!A1+data!B2", "SUM(data:calc!A1)", "const_num*2", "const_str&\"x\"", "SUM(const_arr)", "fml_name", "SUM(ofs)", "local_rng", "SUM(multi)",
  "SUM({1,2,3})", "SUM({1,2;3,4})", "SUM({\"a\",\"b\"})", "SUM(#REF!)", "ERROR.TYPE(#N/A)", "IF(TRUE,#DIV/0!,#VALUE!)",
  "CELL(\"address\",A1)", "INFO(\"osversion\")", "FORMULATEXT(A1)", "ISFORMULA(A1)", "SUBTOTAL(9,A1:A10)", "AGGREGATE(9,6,A1:A10)",
  "REPT(\"ab\",3)", "CODE(\"a\")&CHAR(65)", "VALUE(\"3\")", "T(\"a\")&N(1)", "TYPE(A1)", "NA()", "FIXED(1234.5,1)", "DOLLAR(12.3)",
  "PMT(0.05/12,60,10000)", "NPV(0.1,A1:A5)", "IRR(A1:A5)", "TREND(A1:A5)", "LINEST(A1:A5,A1:A5)", "TRANSPOSE(A1:A3)", "MMULT(A1:B2,A1:B2)",
  "RAND()*0", "RANDBETWEEN(1,1)", "COUNTBLANK(A1:A10)", "COUNTA(A1:A10)", "SMALL(A1:A10,1)", "MODE(A1:A10)", "FREQUENCY(A1:A10,A1:A3)",
  "LOOKUP(2,A1:A5)", "HLOOKUP(1,A1:C2,2)", "MATCH(1,A1:A5,0)", "AREAS(A1:B2)", "ISBLANK(A1)", "ISERROR(A1)", "ISTEXT(A1)", "ISLOGICAL(A1)", "ISREF(A1)",
  "ROMAN(12)", "BIN2DEC(\"101\")", "DEC2HEX(255)", "EOMONTH(DATE(2020,1,1),1)", "WORKDAY(DATE(2020,1,1),5)", "YEARFRAC(DATE(2020,1,1),DATE(2021,1,1))", "DATEDIF(DATE(2020,1,1),DATE(2021,1,1),\"d\")",
  "CONVERT(1,\"m\",\"ft\")", "SUMSQ(A1:A3)", "SUMX2MY2(A1:A3,A1:A3)", "GCD(12,18)+LCM(4,6)", "FACT(5)+COMBIN(5,2)", "CEILING(2.1,1)+FLOOR(2.9,1)", "TRUNC(2.9)+SIGN(-1)",
  "NORMSDIST(0.5)", "NORMINV(0.5,0,1)", "TDIST(1,5,2)", "CHIDIST(1,2)", "FDIST(1,2,3)", "BINOMDIST(1,2,0.5,FALSE)", "POISSON(1,2,FALSE)", "CORREL(A1:A5,A1:A5)", "SLOPE(A1:A5,A1:A5)")
wb$add_data(sheet = "calc", x = data.frame(f = seq_along(f)), start_col = 1)
for (i in seq_along(f)) wb$add_formula(sheet = "calc", x = f[i], dims = wb_dims(rows = i, cols = 3))
wb$add_formula(sheet = "calc", x = "A1:A3*2", dims = "E1:E3", array = TRUE)
wb$add_formula(sheet = "calc", x = "SUM(A1:A10*A1:A10)", dims = "F1", array = TRUE)
wb$workbook$calcPr <- '<calcPr calcId="124519" calcMode="manual" fullCalcOnLoad="1" refMode="R1C1" iterate="1" iterateCount="50" iterateDelta="0.01" fullPrecision="0" calcCompleted="0" calcOnSave="0" concurrentCalc="0"/>'
wb$save(file.path(out, "xls_cov_names_formulas.xlsx"))

# 4 values, strings, styles, long strings, rich text, blank cells, errors, dates
wb <- wb_workbook()
wb$add_worksheet("values")
long <- paste(rep("long string with a few words ", 600), collapse = "")
wb$add_data(x = data.frame(
  l = c(long, strrep("x", 255), strrep("y", 32767), "", " ", "ä ö ü ß € 日本語 😀", "line\nbreak", "tab\there", "<>&\"'"),
  n = c(0, -1.5, 1e300, 1e-300, 123456789012, 0.1 + 0.2, 1/3, NaN, NA),
  i = c(1L, -1L, 65535L, 65536L, 536870911L, 536870912L, -536870912L, -536870913L, NA),
  b = c(TRUE, FALSE, NA, TRUE, FALSE, NA, TRUE, FALSE, NA),
  d = as.Date("2020-01-01") + c(0, 1, 59, 60, 61, 365, 36525, -10, 40000),
  t = as.POSIXct("2020-01-01 12:00:00", tz = "UTC") + c(0, 1, 60, 3600, 86400, 0.5, 100000, 1e6, 1e7),
  stringsAsFactors = FALSE))
wb$add_data(x = data.frame(a = c(-5, 0, 1.5), stringsAsFactors = FALSE), start_col = 8)
wb$add_formula(x = c("#N/A", "1/0"), dims = "J1:J2")
wb$add_formula(x = "1/0", dims = "K1")
wb$add_data(x = matrix(c("#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"), 1), dims = "L1:R1", col_names = FALSE)
wb$add_fill(dims = "A20:C22", color = wb_color("lightblue"))$add_fill(dims = "E20:E22", color = wb_color("#123456"))
wb$add_font(dims = "A20", name = "Times New Roman", size = 20, bold = TRUE, italic = TRUE, underline = "double", strike = TRUE, color = wb_color("red"))
wb$add_font(dims = "A21", name = "Courier New", size = 8, underline = "single", color = wb_color("#7F7F7F"), vert_align = "superscript")
wb$add_font(dims = "A22", name = "Arial", size = 11, vert_align = "subscript", color = wb_color(theme = 4))
wb$add_border(dims = "B20:D24", top_border = "thick", bottom_border = "double", left_border = "dashed", right_border = "dotted", top_color = wb_color("red"), bottom_color = wb_color("blue"))
wb$add_border(dims = "F20:G22", inner_hgrid = "hair", inner_vgrid = "mediumDashDot", left_border = "slantDashDot", right_border = "mediumDashDotDot", top_border = "dashDot", bottom_border = "dashDotDot")
wb$add_cell_style(dims = "A25", horizontal = "center", vertical = "top", wrap_text = TRUE, text_rotation = 45, indent = 2, shrink_to_fit = FALSE)
wb$add_cell_style(dims = "B25", horizontal = "centerContinuous")
wb$add_cell_style(dims = "C25", horizontal = "fill")
wb$add_cell_style(dims = "D25", horizontal = "distributed", vertical = "distributed", text_rotation = 255)
wb$add_cell_style(dims = "E25", horizontal = "justify", vertical = "justify", text_rotation = 90, shrink_to_fit = TRUE)
wb$add_cell_style(dims = "F25", text_rotation = 135, horizontal = "left", indent = 1)
wb$add_cell_style(dims = "G25", horizontal = "right", indent = 3, vertical = "center", reading_order = 2)
wb$add_numfmt(dims = "H1:H3", numfmt = "0.000;[Red]-0.000")
wb$add_numfmt(dims = "A27", numfmt = "#,##0.00 \"EUR\"")$add_numfmt(dims = "A28", numfmt = "[$-409]mmmm d, yyyy")$add_numfmt(dims = "A29", numfmt = "0.0%")
wb$add_numfmt(dims = "A30", numfmt = "# ?/?")$add_numfmt(dims = "A31", numfmt = "0.00E+00")$add_numfmt(dims = "A32", numfmt = "@")$add_numfmt(dims = "A33", numfmt = "[h]:mm:ss")
wb$add_numfmt(dims = "A34", numfmt = "[<100]0.0;[>=100]0;General")$add_numfmt(dims = "A35", numfmt = "yyyy-mm-dd hh:mm AM/PM")
wb$add_data(x = c(1234.567, 44000, 0.256, 0.75, 12345678, 123, 0.5, 44000.75), start_row = 27, dims = "A27")
wb$add_numfmt(dims = "A36", numfmt = "[DBNum1]General")$add_numfmt(dims = "A37", numfmt = "_-* #,##0_-;-* #,##0_-;_-* \"-\"_-;_-@_-")
wb$add_data(x = c(1, 2), dims = "A36")
wb$add_fill(dims = "I20", color = wb_color("#FF0000"), pattern = "darkGrid")$add_fill(dims = "J20", color = wb_color("blue"), pattern = "lightUp")
wb$add_fill(dims = "K20", color = wb_color("blue"), pattern = "gray125")$add_fill(dims = "L20", color = wb_color("#00FF00"), pattern = "darkTrellis")
wb$add_fill(dims = "M20", gradient_fill = '<gradientFill degree="90"><stop position="0"><color rgb="FFFF0000"/></stop><stop position="1"><color rgb="FF0000FF"/></stop></gradientFill>')
wb$add_comment(dims = "A1", comment = "plain note", author = "me")
wb$add_comment(dims = "B2", comment = wb_comment(text = fmt_txt("bold ", bold = TRUE) + fmt_txt("normal ä€"), author = "someone", visible = TRUE, width = 6, height = 3))
wb$add_comment(dims = "C3", comment = strrep("long ", 400), author = "x")
wb$add_data_table(sheet = "values", x = head(iris, 5), dims = "A40", table_style = "TableStyleMedium2", with_filter = TRUE)
wb$add_data(dims = "A50", x = "rich")
wb$add_data(dims = "A51", x = fmt_txt("bold", bold = TRUE) + fmt_txt(" red italic", italic = TRUE, color = wb_color("red")) + fmt_txt(" big", size = 20, font = "Arial") + fmt_txt(" under", underline = TRUE) + fmt_txt(" strike", strike = TRUE))
wb$save(file.path(out, "xls_cov_values.xlsx"))

# 5 images, charts, shapes-free drawing; sparklines are not in xls
wb <- wb_workbook()$add_worksheet("img")
png <- tempfile(fileext = ".png")
grDevices::png(png, 200, 150); plot(1:10); grDevices::dev.off()
jpg <- tempfile(fileext = ".jpg")
grDevices::jpeg(jpg, 200, 150); hist(rnorm(100)); grDevices::dev.off()
wb$add_image(dims = "B2", file = png, width = 3, height = 2, row_offset = 5000, col_offset = 5000)
wb$add_image(dims = "H2", file = jpg, width = 3, height = 2)
wb$add_image(dims = "B20", file = png, width = 3, height = 2, address = "https://www.r-project.org")
wb$add_worksheet("chart")$add_data(x = data.frame(name = letters[1:10], a = c(3, 5, 2, 8, 6, 4, 7, 1, 9, 5), b = (1:10)^2, c = 10:1), col_names = TRUE)
ser <- function(i, col, type = "cat") sprintf('<c:ser><c:idx val="%d"/><c:order val="%d"/><c:tx><c:strRef><c:f>chart!$%s$1</c:f></c:strRef></c:tx>%s%s</c:ser>', i, i, col, if (type == "xy") sprintf('<c:xVal><c:numRef><c:f>chart!$B$2:$B$11</c:f></c:numRef></c:xVal><c:yVal><c:numRef><c:f>chart!$%s$2:$%s$11</c:f></c:numRef></c:yVal>', col, col) else '<c:cat><c:strRef><c:f>chart!$A$2:$A$11</c:f></c:strRef></c:cat>', if (type == "xy") "" else sprintf('<c:val><c:numRef><c:f>chart!$%s$2:$%s$11</c:f></c:numRef></c:val>', col, col))
ax <- '<c:axId val="111"/><c:axId val="222"/>'
axes <- '<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222"/></c:catAx><c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="111"/></c:valAx>'
vaxes <- '<c:valAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222"/></c:valAx><c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111"/></c:valAx>'
plots <- c(
  bar = paste0('<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>', ser(0, "B"), ser(1, "C"), ax, '</c:barChart>', axes),
  barh_stacked = paste0('<c:barChart><c:barDir val="bar"/><c:grouping val="stacked"/>', ser(0, "B"), ser(1, "C"), '<c:overlap val="100"/>', ax, '</c:barChart>', axes),
  line = paste0('<c:lineChart><c:grouping val="standard"/>', ser(0, "B"), ser(1, "D"), '<c:marker val="1"/>', ax, '</c:lineChart>', axes),
  area = paste0('<c:areaChart><c:grouping val="standard"/>', ser(0, "B"), ax, '</c:areaChart>', axes),
  scatter = paste0('<c:scatterChart><c:scatterStyle val="lineMarker"/>', ser(0, "C", "xy"), ser(1, "D", "xy"), ax, '</c:scatterChart>', vaxes),
  pie = paste0('<c:pieChart><c:varyColors val="1"/>', ser(0, "B"), '</c:pieChart>'),
  doughnut = paste0('<c:doughnutChart><c:varyColors val="1"/>', ser(0, "B"), '<c:holeSize val="50"/></c:doughnutChart>'),
  radar = paste0('<c:radarChart><c:radarStyle val="marker"/>', ser(0, "B"), ax, '</c:radarChart>', axes))
for (i in seq_along(plots)) {
  xml <- paste0('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>', names(plots)[i], '</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>', plots[[i]], '</c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/></c:chart></c:chartSpace>')
  wb$add_chart_xml(sheet = "chart", dims = wb_dims(from_row = 14 * (i - 1) + 1, from_col = 8), xml = xml)
}
wb$add_chartsheet("chartsheet")$add_chart_xml(sheet = "chartsheet", xml = xml)
wb$save(file.path(out, "xls_cov_drawings.xlsx"))

# 6 pivot table + external links
wb <- wb_workbook()$add_worksheet("src")$add_data(x = mtcars, row_names = FALSE)
wb$add_worksheet("pivot")
df <- wb_data(wb, sheet = "src")
wb$add_pivot_table(x = df, sheet = "pivot", dims = "A3", rows = c("cyl", "gear"), cols = "am", data = c("mpg", "hp"), filter = "carb")
wb$add_worksheet("pivot2")$add_pivot_table(x = df, sheet = "pivot2", dims = "A3", rows = "carb", data = "wt", fun = "average")
wb$add_worksheet("ext")
wb$add_data(sheet = "ext", x = 1:3)
wb$add_formula(sheet = "ext", x = "SUM('[xls_cov_values.xlsx]values'!H1:H3)", dims = "C1")
wb$add_formula(sheet = "ext", x = "[xls_cov_values.xlsx]values!$A$1", dims = "C2")
wb$save(file.path(out, "xls_cov_pivot_external.xlsx"))
