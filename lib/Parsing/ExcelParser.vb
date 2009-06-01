' ExcelParser.vb
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
Friend Class ExcelParser
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
		MyBase.New(New ExcelTokenizer(input))
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
		MyBase.New(New ExcelTokenizer(input), analyzer)
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

        pattern = New ProductionPattern(CInt(ExcelConstants.FORMULA), "Formula")
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(ExcelConstants.EQ), 0, 1)
        alt.AddProduction(CInt(ExcelConstants.SCALAR_FORMULA), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(ExcelConstants.SCALAR_FORMULA), "ScalarFormula")
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(ExcelConstants.PRIMARY_EXPRESSION), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(ExcelConstants.PRIMARY_EXPRESSION), "PrimaryExpression")
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(ExcelConstants.EXPRESSION), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(ExcelConstants.EXPRESSION), "Expression")
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(ExcelConstants.LOGICAL_EXPRESSION), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(ExcelConstants.LOGICAL_EXPRESSION), "LogicalExpression")
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(ExcelConstants.CONCAT_EXPRESSION), 1, 1)
        alt.AddProduction(CInt(SynteticPatterns.SUBPRODUCTION_1), 0, -1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(ExcelConstants.LOGICAL_OP), "LogicalOp")
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(ExcelConstants.EQ), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(ExcelConstants.GT), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(ExcelConstants.LT), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(ExcelConstants.GTE), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(ExcelConstants.LTE), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(ExcelConstants.NE), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(ExcelConstants.CONCAT_EXPRESSION), "ConcatExpression")
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(ExcelConstants.ADDITIVE_EXPRESSION), 1, 1)
        alt.AddProduction(CInt(SynteticPatterns.SUBPRODUCTION_2), 0, -1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(ExcelConstants.ADDITIVE_EXPRESSION), "AdditiveExpression")
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(ExcelConstants.MULTIPLICATIVE_EXPRESSION), 1, 1)
        alt.AddProduction(CInt(SynteticPatterns.SUBPRODUCTION_3), 0, -1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(ExcelConstants.ADDITIVE_OP), "AdditiveOp")
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(ExcelConstants.ADD), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(ExcelConstants.SUB), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(ExcelConstants.MULTIPLICATIVE_EXPRESSION), "MultiplicativeExpression")
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(ExcelConstants.EXPONENTIATION_EXPRESSION), 1, 1)
        alt.AddProduction(CInt(SynteticPatterns.SUBPRODUCTION_4), 0, -1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(ExcelConstants.MULTIPLICATIVE_OP), "MultiplicativeOp")
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(ExcelConstants.MUL), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(ExcelConstants.DIV), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(ExcelConstants.EXPONENTIATION_EXPRESSION), "ExponentiationExpression")
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(ExcelConstants.PERCENT_EXPRESSION), 1, 1)
        alt.AddProduction(CInt(SynteticPatterns.SUBPRODUCTION_5), 0, -1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(ExcelConstants.PERCENT_EXPRESSION), "PercentExpression")
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(ExcelConstants.UNARY_EXPRESSION), 1, 1)
        alt.AddToken(CInt(ExcelConstants.PERCENT), 0, -1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(ExcelConstants.UNARY_EXPRESSION), "UnaryExpression")
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(SynteticPatterns.SUBPRODUCTION_6), 0, -1)
        alt.AddProduction(CInt(ExcelConstants.BASIC_EXPRESSION), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(ExcelConstants.BASIC_EXPRESSION), "BasicExpression")
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(ExcelConstants.PRIMITIVE), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(ExcelConstants.FUNCTION_CALL), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(ExcelConstants.REFERENCE), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(ExcelConstants.EXPRESSION_GROUP), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(ExcelConstants.REFERENCE), "Reference")
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(ExcelConstants.DEFINED_NAME), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(ExcelConstants.GRID_REFERENCE_EXPRESSION), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(ExcelConstants.GRID_REFERENCE_EXPRESSION), "GridReferenceExpression")
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(ExcelConstants.SHEET_NAME), 0, 1)
        alt.AddProduction(CInt(ExcelConstants.GRID_REFERENCE), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(ExcelConstants.GRID_REFERENCE), "GridReference")
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(ExcelConstants.CELL), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(ExcelConstants.CELL_RANGE), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(ExcelConstants.ROW_RANGE), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(ExcelConstants.COLUMN_RANGE), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(ExcelConstants.FUNCTION_CALL), "FunctionCall")
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(ExcelConstants.FUNCTION_NAME), 1, 1)
        alt.AddProduction(CInt(ExcelConstants.ARGUMENT_LIST), 0, 1)
        alt.AddToken(CInt(ExcelConstants.RIGHT_PAREN), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(ExcelConstants.ARGUMENT_LIST), "ArgumentList")
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(ExcelConstants.EXPRESSION), 1, 1)
        alt.AddProduction(CInt(SynteticPatterns.SUBPRODUCTION_7), 0, -1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(ExcelConstants.EXPRESSION_GROUP), "ExpressionGroup")
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(ExcelConstants.LEFT_PAREN), 1, 1)
        alt.AddProduction(CInt(ExcelConstants.EXPRESSION), 1, 1)
        alt.AddToken(CInt(ExcelConstants.RIGHT_PAREN), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(ExcelConstants.PRIMITIVE), "Primitive")
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(ExcelConstants.NUMBER), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(ExcelConstants.BOOLEAN), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(ExcelConstants.STRING_LITERAL), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(ExcelConstants.ERROR_EXPRESSION), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(ExcelConstants.BOOLEAN), "Boolean")
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(ExcelConstants.TRUE), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(ExcelConstants.FALSE), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(ExcelConstants.ERROR_EXPRESSION), "ErrorExpression")
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(ExcelConstants.DIV_ERROR), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(ExcelConstants.NA_ERROR), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(ExcelConstants.NAME_ERROR), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(ExcelConstants.NULL_ERROR), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(ExcelConstants.REF_ERROR), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(ExcelConstants.VALUE_ERROR), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(ExcelConstants.NUM_ERROR), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(SynteticPatterns.SUBPRODUCTION_1), "Subproduction1")
        pattern.Synthetic = True
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(ExcelConstants.LOGICAL_OP), 1, 1)
        alt.AddProduction(CInt(ExcelConstants.CONCAT_EXPRESSION), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(SynteticPatterns.SUBPRODUCTION_2), "Subproduction2")
        pattern.Synthetic = True
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(ExcelConstants.CONCAT), 1, 1)
        alt.AddProduction(CInt(ExcelConstants.ADDITIVE_EXPRESSION), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(SynteticPatterns.SUBPRODUCTION_3), "Subproduction3")
        pattern.Synthetic = True
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(ExcelConstants.ADDITIVE_OP), 1, 1)
        alt.AddProduction(CInt(ExcelConstants.MULTIPLICATIVE_EXPRESSION), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(SynteticPatterns.SUBPRODUCTION_4), "Subproduction4")
        pattern.Synthetic = True
        alt = New ProductionPatternAlternative()
        alt.AddProduction(CInt(ExcelConstants.MULTIPLICATIVE_OP), 1, 1)
        alt.AddProduction(CInt(ExcelConstants.EXPONENTIATION_EXPRESSION), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(SynteticPatterns.SUBPRODUCTION_5), "Subproduction5")
        pattern.Synthetic = True
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(ExcelConstants.EXP), 1, 1)
        alt.AddProduction(CInt(ExcelConstants.PERCENT_EXPRESSION), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(SynteticPatterns.SUBPRODUCTION_6), "Subproduction6")
        pattern.Synthetic = True
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(ExcelConstants.ADD), 1, 1)
        pattern.AddAlternative(alt)
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(ExcelConstants.SUB), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)

        pattern = New ProductionPattern(CInt(SynteticPatterns.SUBPRODUCTION_7), "Subproduction7")
        pattern.Synthetic = True
        alt = New ProductionPatternAlternative()
        alt.AddToken(CInt(ExcelConstants.ARG_SEPARATOR), 1, 1)
        alt.AddProduction(CInt(ExcelConstants.EXPRESSION), 1, 1)
        pattern.AddAlternative(alt)
        AddPattern(pattern)
    End Sub
End Class
