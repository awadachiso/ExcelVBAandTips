/**
 * Easy-to-read VLookup, HLookup
 * 
 * Create an array using braces
 * The trick is to use commas and semicolons differently.
 * 
 * Keys, Values are NamedRange
 * 
 * Lookup関数の可読性を上げる
 */

'#vertical 
=VLOOKUP(key,CHOOSE({1,2},Keys,Values),2,FALSE)

'#horizontal
=HLOOKUP(key,CHOOSE({1;2},Keys,Values),2,FALSE)