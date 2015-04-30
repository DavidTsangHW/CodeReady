Option Strict Off
Option Explicit On
Module Bas_DataType
	Private Const ModuleName As String = "Bas_DataType"
	
	Public Function IsHex(ByRef HexValue As String) As Boolean
		
		Dim idx As Short
		Dim Temp As String
		
		For idx = 1 To Len(HexValue)
			Temp = Mid(HexValue, idx, 1)
			If Not (Asc(Temp) >= 48 And Asc(Temp) <= 57) And Not (Asc(Temp) >= 65 And Asc(Temp) <= 70) Then
				Exit Function
			End If
		Next 
		
		IsHex = True
		
	End Function
	
	Public Function HexToDec(ByVal HexValue As String) As Double
		
		'Convert Hex  to Dec
		'Max Hex: F (Dec: 15)
		
		Dim DecValue As Double
		
		If Len(HexValue) > 1 Then
			Exit Function
		End If
		
		If IsNumeric(HexValue) = True Then
			HexToDec = CDbl(HexValue)
			Exit Function
		End If
		
		DecValue = Asc(HexValue) - 55
		
		If DecValue > 15 Then
			Exit Function
		End If
		
		HexToDec = DecValue
		
	End Function
	
	Public Function HexCharcode(ByVal HexValue As String) As Double
        Dim charcode As Object = ""
		
		'Convert Hex ASCII code to Dec ASCII code (ASCII 1 - 255)
		
		'Max Value
		'Hex FF
		
		Dim TempValue As String
		TempValue = HexValue
		
		If Len(HexValue) > 2 Then
			Exit Function
		End If
		
		If Len(HexValue) > 1 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object charcode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			charcode = HexToDec(Mid(HexValue, 1, 1)) * 16
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object charcode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		HexCharcode = charcode + HexToDec(Right(HexValue, 1))
		
	End Function
	
	Public Function HexToByteArray(ByVal HexValue As String, ByRef ByteArray() As Byte) As Boolean
		Dim ArraySize As Object
		
		Dim idx As Double
        'Dim ArrayIdx As Double
		
		If IsHex(HexValue) = False Then
			Exit Function
		End If
		
		If Not Len(HexValue) Mod 2 = 0 Then
			Exit Function
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object ArraySize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ArraySize = Len(HexValue) / 2 - 1
		'UPGRADE_WARNING: Couldn't resolve default property of object ArraySize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ReDim ByteArray(ArraySize)
		For idx = 1 To Len(HexValue) Step 2
			ByteArray((idx - 1) / 2) = HexCharcode(Mid(HexValue, idx, 2))
		Next 
		
		HexToByteArray = True
		
	End Function
	
	Public Function HexToString(ByVal HexValue As String) As String
		
		Dim idx As Double
		'UPGRADE_NOTE: Str was upgraded to Str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Str_Renamed As String
		
		If IsHex(HexValue) = False Then
            Exit Function

        End If

        If Not Len(HexValue) Mod 2 = 0 Then
            Exit Function
        End If

        For idx = 1 To Len(HexValue) Step 2
            Str_Renamed = Str_Renamed & HexCharcode(Mid(HexValue, idx, 2))
        Next

        HexToString = Str_Renamed

    End Function

    Public Function DecimalToBinary(ByRef DecimalValue As Integer, ByRef MinimumDigits As Short) As String

        ' Returns a string containing the binary
        ' representation of a positive integer

        Dim result As String
        Dim ExtraDigitsNeeded As Short

        ' Make sure value is not negative
        DecimalValue = System.Math.Abs(DecimalValue)

        ' Construct the binary value

        Do
            result = CStr(DecimalValue Mod 2) & result
            DecimalValue = DecimalValue \ 2
        Loop While DecimalValue > 0

        ' Add leading zeros if needed
        ExtraDigitsNeeded = MinimumDigits - Len(result)
        If ExtraDigitsNeeded > 0 Then
            result = New String("0", ExtraDigitsNeeded) & result
        End If

        DecimalToBinary = result

    End Function

    Public Function HexToBinary(ByRef HexStr As String, Optional ByRef MinimumDigits As Short = 0) As String
        Dim idx As Object

        Dim BinaryString As String
        Dim DecimalValue As Double

        For idx = 1 To Len(HexStr)
            'UPGRADE_WARNING: Couldn't resolve default property of object idx. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            BinaryString = BinaryString & DecimalToBinary(HexToDec(Mid(HexStr, idx, 1)), 4)
        Next

        BinaryString = leading(BinaryString, MinimumDigits, "0")

        HexToBinary = BinaryString

    End Function

    Public Function BinaryToDec(ByRef BinaryString As String) As Integer

        Dim DecimalValue As Integer
        Dim idx As Short

        DecimalValue = 0

        For idx = 0 To Len(BinaryString) - 1

            DecimalValue = DecimalValue + (2 ^ idx) * CDbl(Mid(BinaryString, Len(BinaryString) - idx, 1))

        Next

        BinaryToDec = DecimalValue

    End Function

    'UPGRADE_NOTE: Str was upgraded to Str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Public Function ToHexString(ByRef Str_Renamed As String) As String

        Dim idx As Short
        Dim HexString As String = ""

        If Len(Trim(Str_Renamed)) <= 0 Then
            ToHexString = ""
            Exit Function
        End If

        For idx = 1 To Len(Str_Renamed)

            HexString = HexString & leading(Hex(Asc(Mid(Str_Renamed, idx, 1))), 2, "0")

        Next

        ToHexString = HexString

    End Function
End Module