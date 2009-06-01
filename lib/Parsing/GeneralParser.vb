' GeneralParser.vb
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

'''<remarks>A token stream parser.</remarks>
Friend Class GeneralParser
    Inherits RecursiveDescentParser

    '''<summary>An enumeration with the generated production node
    '''identity constants.</summary>
    Private Enum SynteticPatterns
        [SUBPRODUCTION_1] = 3001
        [SUBPRODUCTION_2] = 3002
        [SUBPRODUCTION_3] = 3003
        [SUBPRODUCTION_4] = 3004
        [SUBPRODUCTION_5] = 3005
        [SUBPRODUCTION_6] = 3006
        [SUBPRODUCTION_7] = 3007
    End Enum

    '''<summary>Creates a new parser.</summary>
    '''
    '''<param name='input'>the input stream to read from</param>
    '''
    '''<exception cref='ParserCreationException'>if the parser
    '''couldn't be initialized correctly</exception>
	Public Sub New(ByVal input As TextReader)
		MyBase.New(New GeneralTokenizer(input))
		CreatePatterns()
	End Sub

    '''<summary>Creates a new parser.</summary>
    '''
    '''<param name='input'>the input stream to read from</param>
    '''
    '''<param name='analyzer'>the analyzer to parse with</param>
    '''
    '''<exception cref='ParserCreationException'>if the parser
    '''couldn't be initialized correctly</exception>
	Public Sub New(ByVal input As TextReader, ByVal analyzer As Analyzer)
		MyBase.New(New GeneralTokenizer(input), analyzer)
		CreatePatterns()
	End Sub

    '''<summary>Initializes the parser by creating all the production
    '''patterns.</summary>
    '''
    '''<exception cref='ParserCreationException'>if the parser
    '''couldn't be initialized correctly</exception>
    Private Sub CreatePatterns()
        Dim pattern As ProductionPattern
        Dim alt As ProductionPatternAlternative

        pattern = New ProductionPattern(CInt(GeneralConstants.FORMULA), "Formula")
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(GeneralConstants.EQ), 0, 1)
        alt.AddProduction(CInt(GeneralConstants.EXPRESSION), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(GeneralConstants.EXPRESSION), "Expression")
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(GeneralConstants.LOGICAL_EXPRESSION), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(GeneralConstants.LOGICAL_EXPRESSION), "LogicalExpression")
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(GeneralConstants.CONCAT_EXPRESSION), 1, 1)
        alt.AddProduction(CInt(SynteticPatterns.SUBPRODUCTION_1), 0, -1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(GeneralConstants.LOGICAL_OP), "LogicalOp")
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(GeneralConstants.EQ), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(GeneralConstants.GT), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(GeneralConstants.LT), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(GeneralConstants.GTE), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(GeneralConstants.LTE), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(GeneralConstants.NE), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(GeneralConstants.CONCAT_EXPRESSION), "ConcatExpression")
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(GeneralConstants.ADDITIVE_EXPRESSION), 1, 1)
        alt.AddProduction(CInt(SynteticPatterns.SUBPRODUCTION_2), 0, -1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(GeneralConstants.ADDITIVE_EXPRESSION), "AdditiveExpression")
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(GeneralConstants.MULTIPLICATIVE_EXPRESSION), 1, 1)
        alt.AddProduction(CInt(SynteticPatterns.SUBPRODUCTION_3), 0, -1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(GeneralConstants.ADDITIVE_OP), "AdditiveOp")
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(GeneralConstants.ADD), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(GeneralConstants.SUB), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(GeneralConstants.MULTIPLICATIVE_EXPRESSION), "MultiplicativeExpression")
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(GeneralConstants.EXPONENTIATION_EXPRESSION), 1, 1)
        alt.AddProduction(CInt(SynteticPatterns.SUBPRODUCTION_4), 0, -1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(GeneralConstants.MULTIPLICATIVE_OP), "MultiplicativeOp")
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(GeneralConstants.MUL), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(GeneralConstants.DIV), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(GeneralConstants.EXPONENTIATION_EXPRESSION), "ExponentiationExpression")
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(GeneralConstants.PERCENT_EXPRESSION), 1, 1)
        alt.AddProduction(CInt(SynteticPatterns.SUBPRODUCTION_5), 0, -1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(GeneralConstants.PERCENT_EXPRESSION), "PercentExpression")
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(GeneralConstants.UNARY_EXPRESSION), 1, 1)
        alt.AddToken(CInt(GeneralConstants.PERCENT), 0, -1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(GeneralConstants.UNARY_EXPRESSION), "UnaryExpression")
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(SynteticPatterns.SUBPRODUCTION_6), 0, -1)
        alt.AddProduction(CInt(GeneralConstants.BASIC_EXPRESSION), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(GeneralConstants.BASIC_EXPRESSION), "BasicExpression")
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(GeneralConstants.PRIMITIVE), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(GeneralConstants.FUNCTION_CALL), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(GeneralConstants.EXPRESSION_GROUP), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(GeneralConstants.FUNCTION_CALL), "FunctionCall")
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(GeneralConstants.FUNCTION_NAME), 1, 1)
        alt.AddProduction(CInt(GeneralConstants.ARGUMENT_LIST), 0, 1)
        alt.AddToken(CInt(GeneralConstants.RIGHT_PAREN), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(GeneralConstants.ARGUMENT_LIST), "ArgumentList")
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(GeneralConstants.EXPRESSION), 1, 1)
        alt.AddProduction(CInt(SynteticPatterns.SUBPRODUCTION_7), 0, -1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(GeneralConstants.EXPRESSION_GROUP), "ExpressionGroup")
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(GeneralConstants.LEFT_PAREN), 1, 1)
        alt.AddProduction(CInt(GeneralConstants.EXPRESSION), 1, 1)
        alt.AddToken(CInt(GeneralConstants.RIGHT_PAREN), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(GeneralConstants.PRIMITIVE), "Primitive")
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(GeneralConstants.NUMBER), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(GeneralConstants.BOOLEAN), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(GeneralConstants.STRING_LITERAL), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(GeneralConstants.DEFINED_NAME), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(GeneralConstants.BOOLEAN), "Boolean")
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(GeneralConstants.TRUE), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(GeneralConstants.FALSE), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(SynteticPatterns.SUBPRODUCTION_1), "Subproduction1")
        pattern.Synthetic = True
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(GeneralConstants.LOGICAL_OP), 1, 1)
        alt.AddProduction(CInt(GeneralConstants.CONCAT_EXPRESSION), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(SynteticPatterns.SUBPRODUCTION_2), "Subproduction2")
        pattern.Synthetic = True
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(GeneralConstants.CONCAT), 1, 1)
        alt.AddProduction(CInt(GeneralConstants.ADDITIVE_EXPRESSION), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(SynteticPatterns.SUBPRODUCTION_3), "Subproduction3")
        pattern.Synthetic = True
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(GeneralConstants.ADDITIVE_OP), 1, 1)
        alt.AddProduction(CInt(GeneralConstants.MULTIPLICATIVE_EXPRESSION), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(SynteticPatterns.SUBPRODUCTION_4), "Subproduction4")
        pattern.Synthetic = True
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(GeneralConstants.MULTIPLICATIVE_OP), 1, 1)
        alt.AddProduction(CInt(GeneralConstants.EXPONENTIATION_EXPRESSION), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(SynteticPatterns.SUBPRODUCTION_5), "Subproduction5")
        pattern.Synthetic = True
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(GeneralConstants.EXP), 1, 1)
        alt.AddProduction(CInt(GeneralConstants.PERCENT_EXPRESSION), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(SynteticPatterns.SUBPRODUCTION_6), "Subproduction6")
        pattern.Synthetic = True
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(GeneralConstants.ADD), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(GeneralConstants.SUB), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(SynteticPatterns.SUBPRODUCTION_7), "Subproduction7")
        pattern.Synthetic = True
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(GeneralConstants.ARG_SEPARATOR), 1, 1)
        alt.AddProduction(CInt(GeneralConstants.EXPRESSION), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)
    End Sub
End Class
