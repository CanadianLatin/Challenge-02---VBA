# VBA-challenge
Module 2 Challenge. VBA Using an excle file with Stocks information in 3 tabs (years).
I did in addition to the bonus some extra work. Auto-fitting the columns and extra information that shows how many records were analyzed.

The only inconvient I got during the process was looping in the different sheets, using “For Each ws In Worksheets/ Next ws” but then I realized, with some help that each Range() and Cells() has to be with ws.Range() or ws.Cells()
