' GeneralTokenizer.vb
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
Friend Class GeneralTokenizer
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

        pattern = New TokenPattern(CInt(GeneralConstants.ADD), "ADD", TokenPattern.PatternType.STRING, "+")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(GeneralConstants.SUB), "SUB", TokenPattern.PatternType.STRING, "-")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(GeneralConstants.MUL), "MUL", TokenPattern.PatternType.STRING, "*")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(GeneralConstants.DIV), "DIV", TokenPattern.PatternType.STRING, "/")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(GeneralConstants.EXP), "EXP", TokenPattern.PatternType.STRING, "^")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(GeneralConstants.CONCAT), "CONCAT", TokenPattern.PatternType.STRING, "&")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(GeneralConstants.LEFT_PAREN), "LEFT_PAREN", TokenPattern.PatternType.STRING, "(")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(GeneralConstants.RIGHT_PAREN), "RIGHT_PAREN", TokenPattern.PatternType.STRING, ")")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(GeneralConstants.PERCENT), "PERCENT", TokenPattern.PatternType.STRING, "%")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(GeneralConstants.WHITESPACE), "WHITESPACE", TokenPattern.PatternType.REGEXP, "\s+")
        pattern.Ignore = True
		AddPattern(pattern)

		' Use the arg separator from the current culture
		Dim argSeparator As String = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator

		pattern = New TokenPattern(CInt(GeneralConstants.ARG_SEPARATOR), "ARG_SEPARATOR", TokenPattern.PatternType.STRING, argSeparator)
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(GeneralConstants.EQ), "EQ", TokenPattern.PatternType.STRING, "=")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(GeneralConstants.LT), "LT", TokenPattern.PatternType.STRING, "<")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(GeneralConstants.GT), "GT", TokenPattern.PatternType.STRING, ">")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(GeneralConstants.LTE), "LTE", TokenPattern.PatternType.STRING, "<=")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(GeneralConstants.GTE), "GTE", TokenPattern.PatternType.STRING, ">=")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(GeneralConstants.NE), "NE", TokenPattern.PatternType.STRING, "<>")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(GeneralConstants.STRING_LITERAL), "STRING_LITERAL", TokenPattern.PatternType.REGEXP, """(""""|[^""])*""")
		AddPattern(pattern)

		' Use the decimal separator from the current culture
		Dim decimalSeparator As String = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator
		Dim numberPattern As String = String.Concat("\d+(\", decimalSeparator, "\d+)?([e][+-]\d{1,3})?")

		pattern = New TokenPattern(CInt(GeneralConstants.NUMBER), "NUMBER", TokenPattern.PatternType.REGEXP, numberPattern)
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(GeneralConstants.TRUE), "TRUE", TokenPattern.PatternType.STRING, "True")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(GeneralConstants.FALSE), "FALSE", TokenPattern.PatternType.STRING, "False")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(GeneralConstants.FUNCTION_NAME), "FUNCTION_NAME", TokenPattern.PatternType.REGEXP, "[a-z][\w]*\(")
        AddPattern(pattern)

        pattern = New TokenPattern(CInt(GeneralConstants.DEFINED_NAME), "DEFINED_NAME", TokenPattern.PatternType.REGEXP, "[_a-z][\w]*")
        AddPattern(pattern)
    End Sub
End Class
