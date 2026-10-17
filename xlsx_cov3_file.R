library(openxlsx2)
out <- if (length(commandArgs(TRUE))) commandArgs(TRUE)[1] else "."

add_cf_xml <- function(wb, sheet, sqref, xml) {
  ws <- wb$worksheets[[sheet]]
  ws$conditionalFormatting <- rbind(
    ws$conditionalFormatting,
    data.frame(sqref = sqref, cf = xml, stringsAsFactors = FALSE)
  )
  wb$worksheets[[sheet]] <- ws
  wb
}

wb <- wb_workbook()
wb$add_dxfs_style("d1", font_color = wb_color("red"), bg_fill = wb_color("yellow"))
wb$add_dxfs_style("d2", bg_fill = wb_color("orange"))

# time periods, one rule per column; dates around today
days <- Sys.Date() + c(-40, -20, -10, -7, -3, -1, 0, 1, 3, 7, 10, 20, 40)
dates <- as.data.frame(replicate(10, days, simplify = FALSE))
names(dates) <- c("today", "yesterday", "tomorrow", "last7Days", "thisWeek", "lastWeek",
                  "nextWeek", "thisMonth", "lastMonth", "nextMonth")
wb$add_worksheet("time")$add_data(x = dates)
tp_formula <- function(type, r) {
  switch(type,
         today = sprintf("FLOOR(%s,1)=TODAY()", r),
         yesterday = sprintf("FLOOR(%s,1)=TODAY()-1", r),
         tomorrow = sprintf("FLOOR(%s,1)=TODAY()+1", r),
         last7Days = sprintf("AND(TODAY()-FLOOR(%s,1)<=6,FLOOR(%s,1)<=TODAY())", r, r),
         thisWeek = sprintf("AND(TODAY()-ROUNDDOWN(%s,0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(%s,0)-TODAY()<=7-WEEKDAY(TODAY()))", r, r),
         lastWeek = sprintf("AND(TODAY()-ROUNDDOWN(%s,0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(%s,0)<(WEEKDAY(TODAY())+7))", r, r),
         nextWeek = sprintf("AND(ROUNDDOWN(%s,0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(%s,0)-TODAY()<(15-WEEKDAY(TODAY())))", r, r),
         thisMonth = sprintf("AND(MONTH(%s)=MONTH(TODAY()),YEAR(%s)=YEAR(TODAY()))", r, r),
         lastMonth = sprintf("AND(MONTH(%s)=MONTH(EDATE(TODAY(),0-1)),YEAR(%s)=YEAR(EDATE(TODAY(),0-1)))", r, r),
         nextMonth = sprintf("AND(MONTH(%s)=MONTH(EDATE(TODAY(),0+1)),YEAR(%s)=YEAR(EDATE(TODAY(),0+1)))", r, r)
  )
}
for (i in seq_along(dates)) {
  col <- int2col(i)
  wb <- add_cf_xml(wb, 1L, paste0(col, "2:", col, "14"), sprintf(
    '<cfRule type="timePeriod" dxfId="%d" priority="%d" timePeriod="%s"><formula>%s</formula></cfRule>',
    i %% 2, i, names(dates)[i], gsub("<", "&lt;", tp_formula(names(dates)[i], paste0(col, "2")), fixed = TRUE)
  ))
}

# average rules
set.seed(1)
vals <- as.data.frame(matrix(round(rnorm(10 * 30, 50, 10), 1), ncol = 10))
names(vals) <- c("above", "below", "aboveEq", "belowEq", "above1", "below1", "above2", "below2", "above3", "below3")
wb$add_worksheet("average")$add_data(x = vals)
avg <- list(
  above   = 'aboveAverage="1"',
  below   = 'aboveAverage="0"',
  aboveEq = 'aboveAverage="1" equalAverage="1"',
  belowEq = 'aboveAverage="0" equalAverage="1"',
  above1  = 'aboveAverage="1" stdDev="1"',
  below1  = 'aboveAverage="0" stdDev="1"',
  above2  = 'aboveAverage="1" stdDev="2"',
  below2  = 'aboveAverage="0" stdDev="2"',
  above3  = 'aboveAverage="1" stdDev="3"',
  below3  = 'aboveAverage="0" stdDev="3"'
)
for (i in seq_along(avg)) {
  col <- int2col(i)
  wb <- add_cf_xml(wb, 2L, paste0(col, "2:", col, "31"), sprintf(
    '<cfRule type="aboveAverage" dxfId="%d" priority="%d" %s/>', i %% 2, i, avg[[i]]
  ))
}

# icon sets and data bars
icons <- c("3Arrows", "3ArrowsGray", "3Flags", "3TrafficLights2", "3Signs", "3Symbols", "3Symbols2",
           "4Arrows", "4ArrowsGray", "4RedToBlack", "4Rating", "4TrafficLights",
           "5ArrowsGray", "5Rating", "5Quarters")
x <- data.frame(matrix(rep(c(-30, -10, 5, 20, 45, 70, 90, 100), length.out = 8 * length(icons)), ncol = length(icons)))
names(x) <- icons
wb$add_worksheet("icons")$add_data(x = x)
for (i in seq_along(icons)) {
  n <- as.integer(substr(icons[i], 1, 1))
  wb$add_conditional_formatting(
    dims = wb_dims(rows = 2:9, cols = i), type = "iconSet",
    rule = round(seq(0, 100, length.out = n + 1)[seq_len(n)]),
    params = list(iconSet = icons[i], reverse = i %% 2 == 0, showValue = i %% 3 != 0)
  )
}

bars <- data.frame(a = c(-50, -20, 0, 30, 80), b = c(1, 5, 10, 20, 40), c = c(-5, 5, -10, 10, 0),
                   d = c(10, 20, 30, 40, 50), e = c(-1, 1, -2, 2, 3))
wb$add_worksheet("bars")$add_data(x = bars)
wb$add_conditional_formatting(dims = "A2:A6", type = "dataBar", style = c("red", "lightgreen"),
                              params = list(gradient = FALSE, border = FALSE, axisPosition = "middle"))
wb$add_conditional_formatting(dims = "B2:B6", type = "dataBar", style = c("blue", "lightblue"),
                              params = list(gradient = TRUE, border = TRUE, showValue = FALSE))
wb$add_conditional_formatting(dims = "C2:C6", type = "dataBar", style = c("orange", "yellow"),
                              params = list(gradient = FALSE, border = TRUE, axisPosition = "none"))
wb$add_conditional_formatting(dims = "D2:D6", type = "dataBar", style = c("purple", "pink"), rule = c(10, 50))
wb$add_conditional_formatting(dims = "E2:E6", type = "dataBar", style = c("darkgreen", "green"),
                              params = list(axisPosition = "automatic"))

# top/bottom variants and stopIfTrue
tb <- as.data.frame(matrix(1:60, ncol = 6))
names(tb) <- c("top3", "bottom3", "top20pct", "bottom20pct", "stop", "between")
wb$add_worksheet("topbottom")$add_data(x = tb)
wb$add_conditional_formatting(dims = "A2:A11", type = "topN", params = list(rank = 3), style = "d1")
wb$add_conditional_formatting(dims = "B2:B11", type = "bottomN", params = list(rank = 3), style = "d2")
wb$add_conditional_formatting(dims = "C2:C11", type = "topN", params = list(rank = 20, percent = TRUE), style = "d1")
wb$add_conditional_formatting(dims = "D2:D11", type = "bottomN", params = list(rank = 20, percent = TRUE), style = "d2")
wb$add_conditional_formatting(dims = "E2:E11", rule = "E2>25", style = "d1")
wb$add_conditional_formatting(dims = "E2:E11", rule = "E2>22", style = "d2")
wb$worksheets[[5]]$conditionalFormatting$cf[5] <- sub('<cfRule ', '<cfRule stopIfTrue="1" ', wb$worksheets[[5]]$conditionalFormatting$cf[5])
wb$add_conditional_formatting(dims = "F2:F11", type = "between", rule = c(3, 7), style = "d1")

wb$save(file.path(out, "xls_cov3_cf.xlsx"))
