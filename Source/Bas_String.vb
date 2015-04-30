Option Strict Off
Option Explicit On
Module Bas_String
	Private Const ModuleName As String = "Bas_String"
	
    Public Function Read_Separated_Text(ByVal Line As String, ByVal Separator As String, ByVal Field As Short) As String

        Dim Start As Short
        Dim TempLine As String
        Dim idx As Short

        TempLine = Line

        For idx = 0 To Field - 1

            Start = InStr(1, TempLine, Separator)

            TempLine = Mid(TempLine, Start + 1)

            If Start = 0 Then
                Read_Separated_Text = ""
                Exit Function
            End If

        Next

        Start = InStr(1, TempLine, Separator)
        If Start > 0 Then
            TempLine = Mid(TempLine, 1, Start - 1)
        End If

        Read_Separated_Text = TempLine

    End Function
	
	Public Function leading(ByVal data As String, ByVal length As Short, ByVal leadingcharacter As String) As String
		
		Dim Counter As Short
		If Len(data) >= length Then
			leading = data
			Exit Function
		End If
		
		Dim Fill As Short
		Fill = CShort(length) - Len(data)
		For Counter = 1 To Fill
			data = leadingcharacter & data
		Next 
		leading = data
		
	End Function
	
    Public Function SliceString(ByVal data As String, ByVal slicelength As Short, ByVal Separator As String) As String
        Dim datalength As Object

        Dim ReturnString As String = ""
        Dim Index As Short

        'UPGRADE_WARNING: Couldn't resolve default property of object datalength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If datalength >= Len(data) Then
            SliceString = data
            Exit Function
        End If

        For Index = 1 To Len(data) - slicelength Step slicelength
            ReturnString = ReturnString & Mid(data, Index, slicelength) & Separator
        Next

        SliceString = ReturnString & Mid(data, Index)

    End Function
	
	'UPGRADE_NOTE: Str was upgraded to Str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function FillString(ByVal Str_Renamed As String, ByVal length As String, ByVal fillcharacter As String) As String
		
		Dim Counter As Short
		If Len(Str_Renamed) >= CDbl(length) Then
			FillString = Str_Renamed
			Exit Function
		End If
		
		Dim Fill As Short
		Fill = CShort(length) - Len(Str_Renamed)
		For Counter = 1 To Fill
			Str_Renamed = Str_Renamed & fillcharacter
		Next 
		FillString = Str_Renamed
		
	End Function
	
	'UPGRADE_NOTE: Str was upgraded to Str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function CenterString(ByVal Str_Renamed As String, ByVal length As Short) As String
		
		Dim Rtn As String
		Dim Fill As Short
		
		Dim Modulus As Short
		
		Rtn = LTrim(Str_Renamed)
		Rtn = Trim(Rtn)
		
		If Len(Rtn) > length Then
			CenterString = Rtn
			Exit Function
		End If
		
		Fill = length - Len(Rtn)
		
		Modulus = Fill Mod 2
		
		Fill = Fill - Modulus
		
		Rtn = Space(Fill / 2) & Rtn & Space(Fill / 2 + Modulus)
		
		CenterString = Rtn
		
	End Function
	
	'UPGRADE_NOTE: Str was upgraded to Str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function RemoveMeta(ByVal Str_Renamed As String) As String
		
		Dim TempStr As String
		Dim Product As String
		'UPGRADE_NOTE: Loc was upgraded to Loc_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Loc_Renamed As Short
		
		TempStr = Str_Renamed
		
		Loc_Renamed = InStr(1, TempStr, "<")
		
		If Loc_Renamed = 0 Then
			Product = TempStr
		End If
		
		Do Until Loc_Renamed <= 0
			
			TempStr = Mid(TempStr, Loc_Renamed + 1)
			
			Loc_Renamed = InStr(1, TempStr, ">")
			
			If Loc_Renamed <= 0 Then
				
				Exit Do
				
			End If
			
			TempStr = Mid(TempStr, Loc_Renamed + 1)

			Loc_Renamed = InStr(1, TempStr, "<")
			
			If Loc_Renamed > 0 Then
				
				Product = Product & Mid(TempStr, 1, Loc_Renamed - 1)
				
			End If
			
		Loop 
		
		RemoveMeta = Product
		
    End Function


End Module