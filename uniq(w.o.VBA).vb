/**
 * Extract Unique Items With Out VBA
 * 
 * A1:A8, B1:B8 contains values,
 *  they are not Table 
 * 
 * 複数の範囲から抽出する
 * VBA不使用
 * 
 */

' col D is list (which named "items")
'  that merge A1:A8 and B1:B8
D1 =IFERROR(INDEX(CHOOSE({1,2},$A$1:$A$8,$B$1:$B$8),IF(ROW()<=COUNTA(A:A),ROW(),ROW()-COUNTA(A:A)),IF(ROW()<=COUNTA(A:A),1,2)),"")

 'COUNTA(A:A) appear over and over again.
 'If you do not want to repeat this expression,
 'you can name the result using NamedRange.

' col E is flag (is Unique?)
E1 =COUNTIF($D$1:D1,D1)=1

' col F is uniques
F1 =IFERROR(INDEX(items,SMALL(IF(flags=TRUE,ROW(items)),ROW())),"")


' ##fill D1,E1,F1 to an appropriate range.##


' namedrange => you can use with data validation
uniques = OFFSET(Sheet1!$F$1,0,0,COUNTIF(Sheet1!flags,TRUE),1)
