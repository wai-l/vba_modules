' Filter Date
Dim filter_date As Date
filter_date = DateAdd("m", -6, Date)
ActiveSheet.Range("E:E").NumberFormat = "dd/mm/yyyy"
ActiveSheet.UsedRange.AutoFilter Field:-5, ,Criteria1:-">=" & CLng(filter_date)