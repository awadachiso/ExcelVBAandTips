/**
 * extract values with key
 * Table : ListObject
 *  that contains Keys, values as columns
 * BaseCell : The first cell of the output range or Namedrange with its cells set
 * OutputRange : that has the maximum expected length
 *
 * 条件を満たすものだけを抽出する
 */

=IFERROR(INDEX(Table[values],SMALL(IF(Table[Keys]=key,ROW(Table[values])-ROW(Table[[#Headers],[values]])),ROW()-ROW(BaseCell)+1)),"")