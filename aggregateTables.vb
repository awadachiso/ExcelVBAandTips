/**
 * Combine multiple tables (records sorted by day)
 *   into a chart for a specific KEY.
 *
 * Table : ListObject
 *  that contains Keys, date as columns
 *  
 *  
 *  Src Table( e.g. tableA,tableB, )
 *  [col]:keys, day1, day2, day3...
 *  
 *  ↓combine
 *  
 *  Dest Table
 *        [column]: A, B,
 *  [row]
 *  day1
 *  day2
 *  day3
 * 
 *  複数のテーブルから特定のキーについてのテーブルに集約したいとき
 */

'# enter below formula into B2 cell
'# row(1) contains SrcTableNames, column("A") contains days 
 =VLOOKUP(KEY,INDIRECT(B$1),MATCH($A2,INDIRECT(B$1&"[#Headers]"),0),FALSE)
 
'# If you use the table named "aggTable"
=VLOOKUP(KEY,INDIRECT(aggTable[[#Headers],[tableA]]),MATCH($A2,INDIRECT(aggTable[[#Headers],[tableA]]&"[#Headers]"),0),FALSE)
