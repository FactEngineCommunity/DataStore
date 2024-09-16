Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports Microsoft.VisualBasic
Imports System.ComponentModel
Imports System.Text.RegularExpressions
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.IO
Imports System.Runtime.Serialization
Imports System.Runtime.InteropServices
Imports System.Globalization
Imports System.Dynamic

Module MyMethodExtensions

    <StructLayout(LayoutKind.Sequential)>
    Private Structure RECT
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure

    ''' <summary>
    ''' Append to an array.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="original"></param>
    ''' <param name="rangeToAppend"></param>
    ''' <returns></returns>
    <Extension()>
    Public Function AppendRange(ByVal original As Array, ByVal rangeToAppend As Array) As Array

        If original.GetType() IsNot rangeToAppend.GetType() Then
            Throw New ArgumentException("The arrays must be of the same type.")
        End If

        Dim resultArray = Array.CreateInstance(original.GetType().GetElementType(), original.Length + rangeToAppend.Length)

        Array.Copy(original, resultArray, original.Length)
        Array.Copy(rangeToAppend, 0, resultArray, original.Length, rangeToAppend.Length)

        Return resultArray
    End Function


    <Extension()>
    Public Function CondenseString(ByVal input As String, ByVal startLength As Integer, ByVal endLength As Integer, ByVal ellipsisCount As Integer) As String
        If input.Length <= startLength + endLength + ellipsisCount Then
            Return input
        End If

        Dim condensedString As String = input.Substring(0, startLength)
        condensedString += New String("."c, ellipsisCount)
        condensedString += input.Substring(input.Length - endLength, endLength)

        Return condensedString
    End Function

    <Extension()>
    Function ToPascalCase(ByVal input As String) As String
        Dim words As String() = input.Split({" "c, "_"c}, StringSplitOptions.None)

        For i As Integer = 0 To words.Length - 1
            words(i) = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(words(i).ToLower())
        Next

        Dim pascalCase As String = String.Join(" ", words)

        Return pascalCase
    End Function


    <Extension()>
    Public Function ToPascalCaseWithSpaces(ByVal input As String) As String
        Try
            Dim pascalWithSpaces As String = Regex.Replace(input, "([a-z])([A-Z])", "$1 $2")
            Return CultureInfo.CurrentCulture.TextInfo.ToTitleCase(pascalWithSpaces)
        Catch ex As Exception
            Return input
        End Try

    End Function

    ''' <summary>
    ''' Limit to use with ExpandoObjects please.
    ''' </summary>
    ''' <param name="expando"></param>
    ''' <param name="fieldName"></param>
    ''' <returns></returns>
    <System.Runtime.CompilerServices.Extension()>
    Public Function HasField(expando As Object, fieldName As String) As Boolean
        Try
            Dim value As Object = CallByName(expando, fieldName, CallType.Get)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    <System.Runtime.CompilerServices.Extension>
    Public Function GetAttributeValue(Of TAttribute As Attribute, TValue)(source As [Enum]) As TValue
        Dim memberInfo As MemberInfo = source.GetType().GetMember(source.ToString()).FirstOrDefault()
        If memberInfo IsNot Nothing Then
            Dim attribute As TAttribute = memberInfo.GetCustomAttribute(Of TAttribute)()
            If attribute IsNot Nothing Then
                Dim valueProperty As PropertyInfo = attribute.GetType().GetProperty("Value")
                If valueProperty IsNot Nothing Then
                    Return DirectCast(valueProperty.GetValue(attribute), TValue)
                End If
            End If
        End If
        Return Nothing
    End Function

    <Extension()>
    Public Sub ReplaceWith(Of T As Class)(ByRef obj As T, other As T)
        Dim size = Marshal.SizeOf(GetType(T))
        Dim ptr = Marshal.AllocHGlobal(size)
        Marshal.StructureToPtr(other, ptr, False)
        Marshal.PtrToStructure(ptr, obj)
        Marshal.FreeHGlobal(ptr)
    End Sub

    ''' <summary>
    ''' Add/Append and item to an array.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="arr"></param>
    ''' <param name="item"></param>
    <Extension()>
    Public Sub Add(Of T)(ByRef arr As T(), item As T)
        Array.Resize(arr, arr.Length + 1)
        arr(arr.Length - 1) = item
    End Sub

    ''' <summary>
    ''' Truncates a string to a max number of characters.
    ''' NB Can also use Strings.Left(str,int)
    ''' </summary>
    ''' <param name="value">The String to truncate</param>
    ''' <param name="maxLength">The maximum length of the string from the left.</param>
    ''' <returns></returns>
    <Extension()>
    Public Function Truncate(ByVal value As String, ByVal maxLength As Integer) As String
        If String.IsNullOrEmpty(value) Then Return value
        Return If(value.Length <= maxLength, value, value.Substring(0, maxLength))
    End Function

    <Extension()>
    Public Function LCase(ByVal value As String) As String
        If String.IsNullOrEmpty(value) Then Return value
        Return Strings.LCase(value)
    End Function


    <Extension()>
    Public Sub RenameKey(Of TKey, TValue)(ByVal dic As IDictionary(Of TKey, TValue), ByVal fromKey As TKey, ByVal toKey As TKey)
        Dim value As TValue = dic(fromKey)
        dic.Remove(fromKey)
        dic(toKey) = value
    End Sub


    <Extension()>
    Public Function AddUnique(Of T)(list As List(Of T), item As T)

        If Not list.Contains(item) Then list.Add(item)
        Return list
    End Function

    <Extension()>
    Public Function MaximumValue(ByVal aiInt As Integer, ByVal aiMaximumValue As Integer) As Integer

        If aiInt > aiMaximumValue Then
            Return aiMaximumValue
        Else
            Return aiInt
        End If


    End Function

    <Extension()>
    Public Function MaximumValue(ByVal aiDbl As Double, ByVal aiMaximumValue As Integer) As Double

        If aiDbl > aiMaximumValue Then
            Return aiMaximumValue
        Else
            Return aiDbl
        End If

    End Function


    <Extension()>
    Public Function CountSubstring(ByVal asString As String, ByVal asSubstring As String) As Integer

        Return asString.Split(asSubstring).Length - 1

    End Function

    <Extension()>
    Public Function GetByDescription(ByRef aiEnum As [Enum], ByVal asDescription As String) As [Enum]

        aiEnum = CType([Enum].Parse(aiEnum.GetType, asDescription), [Enum])
        Return aiEnum

    End Function

    <Extension()>
    Public Function GetEnumValue(Of TEnum)(ByVal value As Integer) As TEnum

        Return CType(System.Enum.ToObject(GetType(TEnum), value), TEnum)
    End Function

    <Extension()>
    Public Function DescriptionAttr(Of T)(ByVal source As T) As String
        Dim fi As FieldInfo = source.[GetType]().GetField(source.ToString())
        Dim attributes As DescriptionAttribute() = CType(fi.GetCustomAttributes(GetType(DescriptionAttribute), False), DescriptionAttribute())

        If attributes IsNot Nothing AndAlso attributes.Length > 0 Then
            Return attributes(0).Description
        Else
            Return source.ToString()
        End If
    End Function


    <Extension()>
    Public Function isBetween(ByRef asglNumber As Single, ByVal aiLowerVal As Integer, ByVal aiUpperVal As Integer) As Boolean

        Return (asglNumber >= aiLowerVal) And (asglNumber <= aiUpperVal)

    End Function

    <Extension()>
    Public Function RemoveDoubleWhiteSpace(ByRef asString As String) As String

        While asString.Contains("  ")
            asString = asString.Replace("  ", " ")
        End While

        Return asString

    End Function

    <Extension()>
    Public Function AppendString(ByRef asString As String, ByVal asStringExtension As String) As String

        asString = asString & asStringExtension
        Return asString

    End Function

    <Extension()>
    Public Function RemoveWhitespace(ByRef asString As String) As String

        Return Regex.Replace(asString, "\s+", "")

    End Function

    <Extension()>
    Public Function AppendLine(ByRef asString As String, ByVal asStringExtension As String) As String

        asString = asString & vbCrLf & asStringExtension
        Return asString

    End Function

    <Extension()>
    Public Function ReplaceFirst(ByRef asString As String, ByVal asFirstString As String, ByVal asReplaceString As String)

        Dim pos As Integer = asString.IndexOf(asFirstString)

        If pos < 0 Then
            Return asString
        Else
            asString = asString.Substring(0, pos) + asReplaceString + asString.Substring(pos + asFirstString.Length)
            Return asString
        End If

    End Function

    <Extension()>
    Public Function AppendDoubleLineBreak(ByRef asString As String, ByVal asStringExtension As String) As String

        asString = asString & vbCrLf & vbCrLf & asStringExtension
        Return asString

    End Function



    <Extension()>
    Public Function IsNumeric(ByRef asString As String) As Boolean

        Dim number As Integer
        Return Int32.TryParse(asString, number)

    End Function

    <Extension()>
    Public Iterator Function Permute(Of T)(ByVal sequence As IEnumerable(Of T)) As IEnumerable(Of IEnumerable(Of T))
        If sequence Is Nothing Then
            Return
        End If

        Dim list = sequence.ToList()

        If Not list.Any() Then
            Yield Enumerable.Empty(Of T)()
        Else
            Dim startingElementIndex = 0

            For Each startingElement In list
                Dim index = startingElementIndex
                Dim remainingItems = list.Where(Function(e, i) i <> index)

                For Each permutationOfRemainder In remainingItems.Permute()
                    Yield permutationOfRemainder.Prepend(startingElement)
                Next

                startingElementIndex += 1
            Next
        End If
    End Function

End Module
