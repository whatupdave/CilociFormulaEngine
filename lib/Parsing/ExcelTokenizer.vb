' ExcelTokenizer.vb
'
' THIS FILE HAS BEEN GENERATED AUTOMATICALLY. DO NOT EDIT!
'
' This library is free software; you can redistribute it and/or
' modify it under the terms of the GNU Lesser General Public License
' as published by the Free Software Foundation; either version 2.1
' of the License, or (at your option) any later version.
'
' This library is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
' Lesser General Public License for more details.
'
' You should have received a copy of the GNU Lesser General Public
' License along with this library; if not, write to the Free
' Software Foundation, Inc., 59 Temple Place, Suite 330, Boston,
' MA 02111-1307, USA.
'
'
' Copyright (c) 2007 Eugene Ciloci

Imports System.IO

Imports PerCederberg.Grammatica.Runtime

'''<remarks>A character stream tokenizer.</remarks>
Friend Class ExcelTokenizer
    Inherits Tokenizer

    '''<summary>Creates a new tokenizer for the specified input
    '''stream.</summary>
    '''
    '''<param name='input'>the input stream to read</param>
    '''
    '''<exception cref='ParserCreationException'>if the tokenizer
    '''couldn't be initialized correctly</exception>
	Public Sub New(ByVal input As TextReader)
		MyBase.New(input, True)
		CreatePatterns()
	End Sub

    '''<summary>Initializes the tokenizer by creating all the token
    '''patterns.</summary>
    '''
    '''<exception cref='ParserCreationException'>if the tokenizer
    '''couldn't be initialized correctly</exception>
    Private Sub CreatePatterns()
        Dim pattern as TokenPattern

        pattern = New TokenPattern(CInt(ExcelConstants.ADD), "ADD", TokenPattern.PatternType.STRING, "+")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(ExcelConstants.SUB), "SUB", TokenPattern.PatternType.STRING, "-")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(ExcelConstants.MUL), "MUL", TokenPattern.PatternType.STRING, "*")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(ExcelConstants.DIV), "DIV", TokenPattern.PatternType.STRING, "/")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(ExcelConstants.EXP), "EXP", TokenPattern.PatternType.STRING, "^")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(ExcelConstants.CONCAT), "CONCAT", TokenPattern.PatternType.STRING, "&")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(ExcelConstants.LEFT_PAREN), "LEFT_PAREN", TokenPattern.PatternType.STRING, "(")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(ExcelConstants.RIGHT_PAREN), "RIGHT_PAREN", TokenPattern.PatternType.STRING, ")")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(ExcelConstants.PERCENT), "PERCENT", TokenPattern.PatternType.STRING, "%")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(ExcelConstants.WHITESPACE), "WHITESPACE", TokenPattern.PatternType.REGEXP, "\s+")
        pattern.Ignore = True
		AddPattern(pattern)

		' Use the arg separator from the current culture
		Dim argSeparator As String = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator

		pattern = New TokenPattern(CInt(ExcelConstants.ARG_SEPARATOR), "ARG_SEPARATOR", TokenPattern.PatternType.STRING, argSeparator)

        AddPattern(pattern)

        pattern = New TokenPattern(CInt(ExcelConstants.EQ), "EQ", TokenPattern.PatternType.STRING, "=")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(ExcelConstants.LT), "LT", TokenPattern.PatternType.STRING, "<")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(ExcelConstants.GT), "GT", TokenPattern.PatternType.STRING, ">")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(ExcelConstants.LTE), "LTE", TokenPattern.PatternType.STRING, "<=")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(ExcelConstants.GTE), "GTE", TokenPattern.PatternType.STRING, ">=")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(ExcelConstants.NE), "NE", TokenPattern.PatternType.STRING, "<>")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(ExcelConstants.STRING_LITERAL), "STRING_LITERAL", TokenPattern.PatternType.REGEXP, """(""""|[^""])*""")
		AddPattern(pattern)

		' Use the decimal separator from the current culture
		Dim decimalSeparator As String = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator
		Dim numberPattern As String = String.Concat("\d+(\", decimalSeparator, "\d+)?([e][+-]\d{1,3})?")

		pattern = New TokenPattern(CInt(ExcelConstants.NUMBER), "NUMBER", TokenPattern.PatternType.REGEXP, numberPattern)
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(ExcelConstants.TRUE), "TRUE", TokenPattern.PatternType.STRING, "True")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(ExcelConstants.FALSE), "FALSE", TokenPattern.PatternType.STRING, "False")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(ExcelConstants.FUNCTION_NAME), "FUNCTION_NAME", TokenPattern.PatternType.REGEXP, "[a-z][\w]*\(")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(ExcelConstants.DIV_ERROR), "DIV_ERROR", TokenPattern.PatternType.STRING, "#DIV/0!")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(ExcelConstants.NA_ERROR), "NA_ERROR", TokenPattern.PatternType.STRING, "#N/A")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(ExcelConstants.NAME_ERROR), "NAME_ERROR", TokenPattern.PatternType.STRING, "#NAME?")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(ExcelConstants.NULL_ERROR), "NULL_ERROR", TokenPattern.PatternType.STRING, "#NULL!")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(ExcelConstants.REF_ERROR), "REF_ERROR", TokenPattern.PatternType.STRING, "#REF!")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(ExcelConstants.VALUE_ERROR), "VALUE_ERROR", TokenPattern.PatternType.STRING, "#VALUE!")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(ExcelConstants.NUM_ERROR), "NUM_ERROR", TokenPattern.PatternType.STRING, "#NUM!")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(ExcelConstants.CELL), "CELL", TokenPattern.PatternType.REGEXP, "\$?[a-z]{1,2}\$?\d{1,5}")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(ExcelConstants.CELL_RANGE), "CELL_RANGE", TokenPattern.PatternType.REGEXP, "\$?[a-z]{1,2}\$?\d{1,5}:\$?[a-z]{1,2}\$?\d{1,5}")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(ExcelConstants.ROW_RANGE), "ROW_RANGE", TokenPattern.PatternType.REGEXP, "\$?\d{1,5}:\$?\d{1,5}")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(ExcelConstants.COLUMN_RANGE), "COLUMN_RANGE", TokenPattern.PatternType.REGEXP, "\$?[a-z]{1,2}:\$?[a-z]{1,2}")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(ExcelConstants.SHEET_NAME), "SHEET_NAME", TokenPattern.PatternType.REGEXP, "[_a-z][\w]*!")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(ExcelConstants.DEFINED_NAME), "DEFINED_NAME", TokenPattern.PatternType.REGEXP, "[_a-z][\w]*")
        AddPattern(pattern)
    End Sub
End Class
