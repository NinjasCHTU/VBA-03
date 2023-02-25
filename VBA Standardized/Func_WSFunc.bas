Attribute VB_Name = "Func_WSFunc"
  '@@@@@@@@@@@@@@@@Dependency -> No
  'For all of these Module
Function VB_Unique(arr, Optional by_col = False, Optional exactly_once = False)
    VB_Unique = WorksheetFunction.Unique(arr, by_col, exactly_once)
End Function
Function VB_Xlookup(lookup_value, lookup_array, return_array, Optional if_not_found, Optional match_mode, Optional search_mode)
    VB_Xlookup = Application.XLookup(lookup_value, lookup_array, return_array, if_not_found, match_mode, search_mode)

End Function


Function VB_IfError(value, value_if_error)
    VB_IfError = WorksheetFunction.IfError(value, value_if_error)
End Function
  
Function VB_Transpose(arr)

    VB_Transpose = WorksheetFunction.Transpose(arr)
End Function

'How to use sort in VBA
'https://excelchamps.com/vba/sort-range/



Function VB_Seq(n, Optional column = 1, Optional start = 1, Optional step = 1)
    res = WorksheetFunction.Sequence(n, column, start, step)
    VB_Seq = res
End Function

Function VB_SortBy(arr, by_array1, Optional sort_order1 = 1)
'1 = ascending
'-1 = descending
    res = WorksheetFunction.SortBy(arr, by_array1, sort_order1)
    VB_SortBy = res
End Function




Function VB_Find(find_text As String, within_text As String, Optional startnum = 1)
    res = WorksheetFunction.Find(find_text, within_text, startnum)
    VB_Find = res
End Function



