Attribute VB_Name = "BubbleSort"
Option Explicit
Public Sub TestBubbleSorting()
    Const DELIMITER As String = ","
    Dim targetSheet As Worksheet
    Set targetSheet = ActiveSheet
    Dim numberOfArrays As Long
    numberOfArrays = targetSheet.Cells(1, 1)
    Dim inputValue As String
    Dim outputValue As String
    Dim targetRow As Long
    Dim index As Long
    Dim rawArray As Variant
    Dim numberArray() As Double

    For targetRow = 2 To numberOfArrays + 1
        inputValue = targetSheet.Cells(targetRow, 1)
        If Replace(inputValue, DELIMITER, vbNullString) = vbNullString Then GoTo NextIteration

        rawArray = GetArrayFromCell(inputValue, DELIMITER)
        
        'Create a sort for alphabetic strings? If so ->
        'Create function to run only if numbers?
        ReDim numberArray(LBound(rawArray) To UBound(rawArray))
        For index = LBound(rawArray) To UBound(rawArray)
            If Not IsNumeric(rawArray(index)) Then GoTo NextIteration
            numberArray(index) = CDbl(rawArray(index))
        Next
        
        BubbleSortNumbers numberArray(), True
        
        outputValue = CreateOutputString(numberArray(), DELIMITER)
        targetSheet.Cells(targetRow, 2) = outputValue
NextIteration:
    Next
End Sub

Private Function GetArrayFromCell(ByVal inputValue As String, ByVal DELIMITER As String) As Variant
    GetArrayFromCell = Split(inputValue, DELIMITER)
End Function

Private Sub BubbleSortNumbers(ByRef numberArray() As Double, Optional ByVal sortAscending As Boolean = True)
    Dim index As Long
    Dim isChanged As Boolean
    Dim firstPosition As Long
    Dim lastPosition As Long
    firstPosition = LBound(numberArray)
    lastPosition = UBound(numberArray) - 1
    
    If sortAscending Then
        Do
            isChanged = False
            For index = firstPosition To lastPosition
                If numberArray(index) > numberArray(index + 1) Then
                    isChanged = True
                    SwapElements numberArray(), index, index + 1
                End If
            Next index
            lastPosition = lastPosition - 1
        Loop While isChanged
    Else
         Do
            isChanged = False
            For index = firstPosition To lastPosition
                If numberArray(index) < numberArray(index + 1) Then
                    isChanged = True
                    SwapElements numberArray(), index, index + 1

                End If
            Next index
            lastPosition = lastPosition - 1
        Loop While isChanged
    End If
End Sub

Private Sub SwapElements(ByRef numberArray() As Double, ByVal i As Long, ByVal j As Long)
    Dim temporaryHolder As Double
    temporaryHolder = numberArray(i)
    numberArray(i) = numberArray(j)
    numberArray(j) = temporaryHolder
End Sub

Private Function CreateOutputString(ByVal numberArray As Variant, ByVal DELIMITER As String) As String
    Dim index As Long
    For index = LBound(numberArray) To UBound(numberArray) - 1
            CreateOutputString = CreateOutputString & numberArray(index) & DELIMITER
    Next
    CreateOutputString = CreateOutputString & numberArray(UBound(numberArray))
End Function


