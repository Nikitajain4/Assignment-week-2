{\rtf1\ansi\ansicpg1252\cocoartf2708
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\margl1440\margr1440\vieww11520\viewh8400\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf0 Sub ticker()\
\
'codes to work in all the worksheets in the workbook\
\
For Each ws In Worksheets\
\
    ' Set an initial variable for holding the ticker,closing_value,stockvolume_total,ticker_table_row\
    \
  Dim ticker As String\
  Dim opening_value As Double\
  Dim closing_value As Double\
  Dim stockvolume_Total As Double\
  Dim ticker_Table_Row As Integer\
  \
   'to print the cell names in the respective rows\
   \
   ws.Cells(1, 8).Value = "ticker"\
   ws.Cells(1, 9).Value = "yearly change"\
   ws.Cells(1, 10).Value = "percentage change"\
   ws.Cells(1, 11).Value = "total stock volume"\
   \
    ' Keep track of the location for each ticker,ticker_table_row and stockvolume_total\
   stockvolume_Total = 0\
   ticker_Table_Row = 2\
   firstrow_ticker = 2\
\
  \
  'to read the records till the last row\
  \
     lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row\
      \
      ' Loop through all ticker\
      \
  For i = 2 To lastrow\
      'to add the stockvolume of a particular ticker\
      \
         stockvolume_Total = stockvolume_Total + ws.Cells(i, 7).Value\
      \
          If firstrow_ticker = 2 Then\
          \
       'to read the opening value\
                opening_value = ws.Cells(i, 3).Value\
       \
          End If\
      ' Check if we are still within the same ticker name, if it is not...\
\
          If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then\
\
       'to print the ticker when it changes\
      \
                ticker = ws.Cells(i, 1).Value\
                \
        'to print the last closing value of ticker before it changes\
                \
                closing_value = ws.Cells(i, 6).Value\
                firstrow_ticker = 2\
                \
        ' Print the ticker value in ticker table,yearly change,percentage change\
            \
                  ws.Range("h" & ticker_Table_Row).Value = ticker\
      \
                  ws.Range("i" & ticker_Table_Row).Value = opening_value - closing_value\
   \
                   If opening_value - closing_value > 0 Then\
                   \
        'to change the color of the cell based on positive and negative values\
               \
                          ws.Range("i" & ticker_Table_Row).Interior.ColorIndex = 4\
         \
                   Else\
         \
                          ws.Range("i" & ticker_Table_Row).Interior.ColorIndex = 3\
        \
                   End If\
    \
    \
        'to calculate the percentage change and formatting it with percent\
    \
           ws.Range("j" & ticker_Table_Row).Value = ((closing_value - opening_value) / opening_value)\
         \
           ws.Range("j" & ticker_Table_Row).NumberFormat = "0.00%"\
           \
         'to add the total stock volume of a particular ticker\
    \
           ws.Range("k" & ticker_Table_Row).Value = stockvolume_Total\
\
           ticker_Table_Row = ticker_Table_Row + 1\
      \
           stockvolume_Total = 0\
           \
    Else\
        \
                firstrow_ticker = firstrow_ticker + 1\
           \
     End If\
\
 Next i\
 \
        'to find the percentage increase by using the max in an array\
 \
            ws.Cells(2, 13).Value = "Greatest Percentage Increase"\
 \
            Max_increase = WorksheetFunction.Max(ws.Range("j:j"))\
 \
 \
        'to print the values of percentage increase after finding the column no.\
 \
 \
            max_increase_location = WorksheetFunction.Match(Max_increase, ws.Range("j:j"), 0)\
 \
            ws.Cells(2, 14).Value = ws.Cells(max_increase_location, 8)\
    \
            ws.Cells(2, 15).Value = Max_increase\
 \
           ws.Cells(2, 15).NumberFormat = "0.00%"\
 \
 \
  \
         'to find the percentage decreaseby using the min in an array\
  \
           ws.Cells(3, 13).Value = "Greatest Percentage Decrease"\
 \
           Max_decrease = WorksheetFunction.Min(ws.Range("j:j"))\
 \
 \
         'to print the values of percentage decrease after finding the column no.\
 \
 \
          max_decrease_location = WorksheetFunction.Match(Max_decrease, ws.Range("j:j"), 0)\
 \
          ws.Cells(3, 14).Value = ws.Cells(max_decrease_location, 8)\
    \
          ws.Cells(3, 15).Value = Max_decrease\
 \
          ws.Cells(3, 15).NumberFormat = "0.00%"\
 \
 \
         'to find the greatest total volume in an array\
  \
         ws.Cells(4, 13).Value = "Greatest Total Volume"\
 \
        Max_volume = WorksheetFunction.Max(ws.Range("k:k"))\
 \
 \
          'to print the value of greatest total volume after finding the column no.\
 \
 \
       max_volume_location = WorksheetFunction.Match(Max_volume, ws.Range("k:k"), 0)\
 \
       ws.Cells(4, 14).Value = ws.Cells(max_volume_location, 8)\
    \
       ws.Cells(4, 15).Value = Max_volume\
       \
 \
         'to go to next worksheet\
        \
  Next ws\
 \
 End Sub\
\
\
\
}