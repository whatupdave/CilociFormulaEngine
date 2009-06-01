' ExcelAnalyzer.vb
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

Imports PerCederberg.Grammatica.Runtime

'''<remarks>A class providing callback methods for the
'''parser.</remarks>
Friend MustInherit Class ExcelAnalyzer
    Inherits Analyzer

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overrides Sub Enter(ByVal node As Node)
        Select Case node.Id
        Case ExcelConstants.ADD
            EnterAdd(CType(node,Token))

        Case ExcelConstants.SUB
            EnterSub(CType(node,Token))

        Case ExcelConstants.MUL
            EnterMul(CType(node,Token))

        Case ExcelConstants.DIV
            EnterDiv(CType(node,Token))

        Case ExcelConstants.EXP
            EnterExp(CType(node,Token))

        Case ExcelConstants.CONCAT
            EnterConcat(CType(node,Token))

        Case ExcelConstants.LEFT_PAREN
            EnterLeftParen(CType(node,Token))

        Case ExcelConstants.RIGHT_PAREN
            EnterRightParen(CType(node,Token))

        Case ExcelConstants.PERCENT
            EnterPercent(CType(node,Token))

        Case ExcelConstants.ARG_SEPARATOR
            EnterArgSeparator(CType(node,Token))

        Case ExcelConstants.EQ
            EnterEq(CType(node,Token))

        Case ExcelConstants.LT
            EnterLt(CType(node,Token))

        Case ExcelConstants.GT
            EnterGt(CType(node,Token))

        Case ExcelConstants.LTE
            EnterLte(CType(node,Token))

        Case ExcelConstants.GTE
            EnterGte(CType(node,Token))

        Case ExcelConstants.NE
            EnterNe(CType(node,Token))

        Case ExcelConstants.STRING_LITERAL
            EnterStringLiteral(CType(node,Token))

        Case ExcelConstants.NUMBER
            EnterNumber(CType(node,Token))

        Case ExcelConstants.TRUE
            EnterTrue(CType(node,Token))

        Case ExcelConstants.FALSE
            EnterFalse(CType(node,Token))

        Case ExcelConstants.FUNCTION_NAME
            EnterFunctionName(CType(node,Token))

        Case ExcelConstants.DIV_ERROR
            EnterDivError(CType(node,Token))

        Case ExcelConstants.NA_ERROR
            EnterNaError(CType(node,Token))

        Case ExcelConstants.NAME_ERROR
            EnterNameError(CType(node,Token))

        Case ExcelConstants.NULL_ERROR
            EnterNullError(CType(node,Token))

        Case ExcelConstants.REF_ERROR
            EnterRefError(CType(node,Token))

        Case ExcelConstants.VALUE_ERROR
            EnterValueError(CType(node,Token))

        Case ExcelConstants.NUM_ERROR
            EnterNumError(CType(node,Token))

        Case ExcelConstants.CELL
            EnterCell(CType(node,Token))

        Case ExcelConstants.CELL_RANGE
            EnterCellRange(CType(node,Token))

        Case ExcelConstants.ROW_RANGE
            EnterRowRange(CType(node,Token))

        Case ExcelConstants.COLUMN_RANGE
            EnterColumnRange(CType(node,Token))

        Case ExcelConstants.SHEET_NAME
            EnterSheetName(CType(node,Token))

        Case ExcelConstants.DEFINED_NAME
            EnterDefinedName(CType(node,Token))

        Case ExcelConstants.FORMULA
            EnterFormula(CType(node,Production))

        Case ExcelConstants.SCALAR_FORMULA
            EnterScalarFormula(CType(node,Production))

        Case ExcelConstants.PRIMARY_EXPRESSION
            EnterPrimaryExpression(CType(node,Production))

        Case ExcelConstants.EXPRESSION
            EnterExpression(CType(node,Production))

        Case ExcelConstants.LOGICAL_EXPRESSION
            EnterLogicalExpression(CType(node,Production))

        Case ExcelConstants.LOGICAL_OP
            EnterLogicalOp(CType(node,Production))

        Case ExcelConstants.CONCAT_EXPRESSION
            EnterConcatExpression(CType(node,Production))

        Case ExcelConstants.ADDITIVE_EXPRESSION
            EnterAdditiveExpression(CType(node,Production))

        Case ExcelConstants.ADDITIVE_OP
            EnterAdditiveOp(CType(node,Production))

        Case ExcelConstants.MULTIPLICATIVE_EXPRESSION
            EnterMultiplicativeExpression(CType(node,Production))

        Case ExcelConstants.MULTIPLICATIVE_OP
            EnterMultiplicativeOp(CType(node,Production))

        Case ExcelConstants.EXPONENTIATION_EXPRESSION
            EnterExponentiationExpression(CType(node,Production))

        Case ExcelConstants.PERCENT_EXPRESSION
            EnterPercentExpression(CType(node,Production))

        Case ExcelConstants.UNARY_EXPRESSION
            EnterUnaryExpression(CType(node,Production))

        Case ExcelConstants.BASIC_EXPRESSION
            EnterBasicExpression(CType(node,Production))

        Case ExcelConstants.REFERENCE
            EnterReference(CType(node,Production))

        Case ExcelConstants.GRID_REFERENCE_EXPRESSION
            EnterGridReferenceExpression(CType(node,Production))

        Case ExcelConstants.GRID_REFERENCE
            EnterGridReference(CType(node,Production))

        Case ExcelConstants.FUNCTION_CALL
            EnterFunctionCall(CType(node,Production))

        Case ExcelConstants.ARGUMENT_LIST
            EnterArgumentList(CType(node,Production))

        Case ExcelConstants.EXPRESSION_GROUP
            EnterExpressionGroup(CType(node,Production))

        Case ExcelConstants.PRIMITIVE
            EnterPrimitive(CType(node,Production))

        Case ExcelConstants.BOOLEAN
            EnterBoolean(CType(node,Production))

        Case ExcelConstants.ERROR_EXPRESSION
            EnterErrorExpression(CType(node,Production))

        End Select
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overrides Function [Exit](ByVal node As Node) As Node
        Select Case node.Id
        Case ExcelConstants.ADD
            return ExitAdd(CType(node,Token))

        Case ExcelConstants.SUB
            return ExitSub(CType(node,Token))

        Case ExcelConstants.MUL
            return ExitMul(CType(node,Token))

        Case ExcelConstants.DIV
            return ExitDiv(CType(node,Token))

        Case ExcelConstants.EXP
            return ExitExp(CType(node,Token))

        Case ExcelConstants.CONCAT
            return ExitConcat(CType(node,Token))

        Case ExcelConstants.LEFT_PAREN
            return ExitLeftParen(CType(node,Token))

        Case ExcelConstants.RIGHT_PAREN
            return ExitRightParen(CType(node,Token))

        Case ExcelConstants.PERCENT
            return ExitPercent(CType(node,Token))

        Case ExcelConstants.ARG_SEPARATOR
            return ExitArgSeparator(CType(node,Token))

        Case ExcelConstants.EQ
            return ExitEq(CType(node,Token))

        Case ExcelConstants.LT
            return ExitLt(CType(node,Token))

        Case ExcelConstants.GT
            return ExitGt(CType(node,Token))

        Case ExcelConstants.LTE
            return ExitLte(CType(node,Token))

        Case ExcelConstants.GTE
            return ExitGte(CType(node,Token))

        Case ExcelConstants.NE
            return ExitNe(CType(node,Token))

        Case ExcelConstants.STRING_LITERAL
            return ExitStringLiteral(CType(node,Token))

        Case ExcelConstants.NUMBER
            return ExitNumber(CType(node,Token))

        Case ExcelConstants.TRUE
            return ExitTrue(CType(node,Token))

        Case ExcelConstants.FALSE
            return ExitFalse(CType(node,Token))

        Case ExcelConstants.FUNCTION_NAME
            return ExitFunctionName(CType(node,Token))

        Case ExcelConstants.DIV_ERROR
            return ExitDivError(CType(node,Token))

        Case ExcelConstants.NA_ERROR
            return ExitNaError(CType(node,Token))

        Case ExcelConstants.NAME_ERROR
            return ExitNameError(CType(node,Token))

        Case ExcelConstants.NULL_ERROR
            return ExitNullError(CType(node,Token))

        Case ExcelConstants.REF_ERROR
            return ExitRefError(CType(node,Token))

        Case ExcelConstants.VALUE_ERROR
            return ExitValueError(CType(node,Token))

        Case ExcelConstants.NUM_ERROR
            return ExitNumError(CType(node,Token))

        Case ExcelConstants.CELL
            return ExitCell(CType(node,Token))

        Case ExcelConstants.CELL_RANGE
            return ExitCellRange(CType(node,Token))

        Case ExcelConstants.ROW_RANGE
            return ExitRowRange(CType(node,Token))

        Case ExcelConstants.COLUMN_RANGE
            return ExitColumnRange(CType(node,Token))

        Case ExcelConstants.SHEET_NAME
            return ExitSheetName(CType(node,Token))

        Case ExcelConstants.DEFINED_NAME
            return ExitDefinedName(CType(node,Token))

        Case ExcelConstants.FORMULA
            return ExitFormula(CType(node,Production))

        Case ExcelConstants.SCALAR_FORMULA
            return ExitScalarFormula(CType(node,Production))

        Case ExcelConstants.PRIMARY_EXPRESSION
            return ExitPrimaryExpression(CType(node,Production))

        Case ExcelConstants.EXPRESSION
            return ExitExpression(CType(node,Production))

        Case ExcelConstants.LOGICAL_EXPRESSION
            return ExitLogicalExpression(CType(node,Production))

        Case ExcelConstants.LOGICAL_OP
            return ExitLogicalOp(CType(node,Production))

        Case ExcelConstants.CONCAT_EXPRESSION
            return ExitConcatExpression(CType(node,Production))

        Case ExcelConstants.ADDITIVE_EXPRESSION
            return ExitAdditiveExpression(CType(node,Production))

        Case ExcelConstants.ADDITIVE_OP
            return ExitAdditiveOp(CType(node,Production))

        Case ExcelConstants.MULTIPLICATIVE_EXPRESSION
            return ExitMultiplicativeExpression(CType(node,Production))

        Case ExcelConstants.MULTIPLICATIVE_OP
            return ExitMultiplicativeOp(CType(node,Production))

        Case ExcelConstants.EXPONENTIATION_EXPRESSION
            return ExitExponentiationExpression(CType(node,Production))

        Case ExcelConstants.PERCENT_EXPRESSION
            return ExitPercentExpression(CType(node,Production))

        Case ExcelConstants.UNARY_EXPRESSION
            return ExitUnaryExpression(CType(node,Production))

        Case ExcelConstants.BASIC_EXPRESSION
            return ExitBasicExpression(CType(node,Production))

        Case ExcelConstants.REFERENCE
            return ExitReference(CType(node,Production))

        Case ExcelConstants.GRID_REFERENCE_EXPRESSION
            return ExitGridReferenceExpression(CType(node,Production))

        Case ExcelConstants.GRID_REFERENCE
            return ExitGridReference(CType(node,Production))

        Case ExcelConstants.FUNCTION_CALL
            return ExitFunctionCall(CType(node,Production))

        Case ExcelConstants.ARGUMENT_LIST
            return ExitArgumentList(CType(node,Production))

        Case ExcelConstants.EXPRESSION_GROUP
            return ExitExpressionGroup(CType(node,Production))

        Case ExcelConstants.PRIMITIVE
            return ExitPrimitive(CType(node,Production))

        Case ExcelConstants.BOOLEAN
            return ExitBoolean(CType(node,Production))

        Case ExcelConstants.ERROR_EXPRESSION
            return ExitErrorExpression(CType(node,Production))

        End Select
        Return node
    End Function

    '''<summary>Called when adding a child to a parse tree
    '''node.</summary>
    '''
    '''<param name='node'>the parent node</param>
    '''<param name='child'>the child node, or null</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overrides Sub Child(ByVal node As Production, ByVal child As Node)
        Select Case node.Id
        Case ExcelConstants.FORMULA
            ChildFormula(node, child)

        Case ExcelConstants.SCALAR_FORMULA
            ChildScalarFormula(node, child)

        Case ExcelConstants.PRIMARY_EXPRESSION
            ChildPrimaryExpression(node, child)

        Case ExcelConstants.EXPRESSION
            ChildExpression(node, child)

        Case ExcelConstants.LOGICAL_EXPRESSION
            ChildLogicalExpression(node, child)

        Case ExcelConstants.LOGICAL_OP
            ChildLogicalOp(node, child)

        Case ExcelConstants.CONCAT_EXPRESSION
            ChildConcatExpression(node, child)

        Case ExcelConstants.ADDITIVE_EXPRESSION
            ChildAdditiveExpression(node, child)

        Case ExcelConstants.ADDITIVE_OP
            ChildAdditiveOp(node, child)

        Case ExcelConstants.MULTIPLICATIVE_EXPRESSION
            ChildMultiplicativeExpression(node, child)

        Case ExcelConstants.MULTIPLICATIVE_OP
            ChildMultiplicativeOp(node, child)

        Case ExcelConstants.EXPONENTIATION_EXPRESSION
            ChildExponentiationExpression(node, child)

        Case ExcelConstants.PERCENT_EXPRESSION
            ChildPercentExpression(node, child)

        Case ExcelConstants.UNARY_EXPRESSION
            ChildUnaryExpression(node, child)

        Case ExcelConstants.BASIC_EXPRESSION
            ChildBasicExpression(node, child)

        Case ExcelConstants.REFERENCE
            ChildReference(node, child)

        Case ExcelConstants.GRID_REFERENCE_EXPRESSION
            ChildGridReferenceExpression(node, child)

        Case ExcelConstants.GRID_REFERENCE
            ChildGridReference(node, child)

        Case ExcelConstants.FUNCTION_CALL
            ChildFunctionCall(node, child)

        Case ExcelConstants.ARGUMENT_LIST
            ChildArgumentList(node, child)

        Case ExcelConstants.EXPRESSION_GROUP
            ChildExpressionGroup(node, child)

        Case ExcelConstants.PRIMITIVE
            ChildPrimitive(node, child)

        Case ExcelConstants.BOOLEAN
            ChildBoolean(node, child)

        Case ExcelConstants.ERROR_EXPRESSION
            ChildErrorExpression(node, child)

        End Select
    End Sub

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterAdd(ByVal node As Token)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitAdd(ByVal node As Token) As Node
        Return node
    End Function

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterSub(ByVal node As Token)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitSub(ByVal node As Token) As Node
        Return node
    End Function

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterMul(ByVal node As Token)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitMul(ByVal node As Token) As Node
        Return node
    End Function

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterDiv(ByVal node As Token)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitDiv(ByVal node As Token) As Node
        Return node
    End Function

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterExp(ByVal node As Token)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitExp(ByVal node As Token) As Node
        Return node
    End Function

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterConcat(ByVal node As Token)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitConcat(ByVal node As Token) As Node
        Return node
    End Function

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterLeftParen(ByVal node As Token)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitLeftParen(ByVal node As Token) As Node
        Return node
    End Function

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterRightParen(ByVal node As Token)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitRightParen(ByVal node As Token) As Node
        Return node
    End Function

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterPercent(ByVal node As Token)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitPercent(ByVal node As Token) As Node
        Return node
    End Function

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterArgSeparator(ByVal node As Token)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitArgSeparator(ByVal node As Token) As Node
        Return node
    End Function

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterEq(ByVal node As Token)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitEq(ByVal node As Token) As Node
        Return node
    End Function

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterLt(ByVal node As Token)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitLt(ByVal node As Token) As Node
        Return node
    End Function

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterGt(ByVal node As Token)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitGt(ByVal node As Token) As Node
        Return node
    End Function

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterLte(ByVal node As Token)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitLte(ByVal node As Token) As Node
        Return node
    End Function

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterGte(ByVal node As Token)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitGte(ByVal node As Token) As Node
        Return node
    End Function

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterNe(ByVal node As Token)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitNe(ByVal node As Token) As Node
        Return node
    End Function

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterStringLiteral(ByVal node As Token)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitStringLiteral(ByVal node As Token) As Node
        Return node
    End Function

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterNumber(ByVal node As Token)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitNumber(ByVal node As Token) As Node
        Return node
    End Function

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterTrue(ByVal node As Token)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitTrue(ByVal node As Token) As Node
        Return node
    End Function

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterFalse(ByVal node As Token)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitFalse(ByVal node As Token) As Node
        Return node
    End Function

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterFunctionName(ByVal node As Token)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitFunctionName(ByVal node As Token) As Node
        Return node
    End Function

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterDivError(ByVal node As Token)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitDivError(ByVal node As Token) As Node
        Return node
    End Function

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterNaError(ByVal node As Token)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitNaError(ByVal node As Token) As Node
        Return node
    End Function

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterNameError(ByVal node As Token)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitNameError(ByVal node As Token) As Node
        Return node
    End Function

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterNullError(ByVal node As Token)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitNullError(ByVal node As Token) As Node
        Return node
    End Function

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterRefError(ByVal node As Token)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitRefError(ByVal node As Token) As Node
        Return node
    End Function

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterValueError(ByVal node As Token)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitValueError(ByVal node As Token) As Node
        Return node
    End Function

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterNumError(ByVal node As Token)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitNumError(ByVal node As Token) As Node
        Return node
    End Function

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterCell(ByVal node As Token)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitCell(ByVal node As Token) As Node
        Return node
    End Function

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterCellRange(ByVal node As Token)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitCellRange(ByVal node As Token) As Node
        Return node
    End Function

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterRowRange(ByVal node As Token)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitRowRange(ByVal node As Token) As Node
        Return node
    End Function

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterColumnRange(ByVal node As Token)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitColumnRange(ByVal node As Token) As Node
        Return node
    End Function

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterSheetName(ByVal node As Token)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitSheetName(ByVal node As Token) As Node
        Return node
    End Function

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterDefinedName(ByVal node As Token)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitDefinedName(ByVal node As Token) As Node
        Return node
    End Function

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterFormula(ByVal node As Production)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitFormula(ByVal node As Production) As Node
        Return node
    End Function

    '''<summary>Called when adding a child to a parse tree
    '''node.</summary>
    '''
    '''<param name='node'>the parent node</param>
    '''<param name='child'>the child node, or null</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub ChildFormula(ByVal node As Production, ByVal child As Node)
        node.AddChild(child)
    End Sub

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterScalarFormula(ByVal node As Production)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitScalarFormula(ByVal node As Production) As Node
        Return node
    End Function

    '''<summary>Called when adding a child to a parse tree
    '''node.</summary>
    '''
    '''<param name='node'>the parent node</param>
    '''<param name='child'>the child node, or null</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub ChildScalarFormula(ByVal node As Production, ByVal child As Node)
        node.AddChild(child)
    End Sub

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterPrimaryExpression(ByVal node As Production)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitPrimaryExpression(ByVal node As Production) As Node
        Return node
    End Function

    '''<summary>Called when adding a child to a parse tree
    '''node.</summary>
    '''
    '''<param name='node'>the parent node</param>
    '''<param name='child'>the child node, or null</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub ChildPrimaryExpression(ByVal node As Production, ByVal child As Node)
        node.AddChild(child)
    End Sub

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterExpression(ByVal node As Production)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitExpression(ByVal node As Production) As Node
        Return node
    End Function

    '''<summary>Called when adding a child to a parse tree
    '''node.</summary>
    '''
    '''<param name='node'>the parent node</param>
    '''<param name='child'>the child node, or null</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub ChildExpression(ByVal node As Production, ByVal child As Node)
        node.AddChild(child)
    End Sub

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterLogicalExpression(ByVal node As Production)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitLogicalExpression(ByVal node As Production) As Node
        Return node
    End Function

    '''<summary>Called when adding a child to a parse tree
    '''node.</summary>
    '''
    '''<param name='node'>the parent node</param>
    '''<param name='child'>the child node, or null</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub ChildLogicalExpression(ByVal node As Production, ByVal child As Node)
        node.AddChild(child)
    End Sub

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterLogicalOp(ByVal node As Production)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitLogicalOp(ByVal node As Production) As Node
        Return node
    End Function

    '''<summary>Called when adding a child to a parse tree
    '''node.</summary>
    '''
    '''<param name='node'>the parent node</param>
    '''<param name='child'>the child node, or null</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub ChildLogicalOp(ByVal node As Production, ByVal child As Node)
        node.AddChild(child)
    End Sub

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterConcatExpression(ByVal node As Production)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitConcatExpression(ByVal node As Production) As Node
        Return node
    End Function

    '''<summary>Called when adding a child to a parse tree
    '''node.</summary>
    '''
    '''<param name='node'>the parent node</param>
    '''<param name='child'>the child node, or null</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub ChildConcatExpression(ByVal node As Production, ByVal child As Node)
        node.AddChild(child)
    End Sub

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterAdditiveExpression(ByVal node As Production)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitAdditiveExpression(ByVal node As Production) As Node
        Return node
    End Function

    '''<summary>Called when adding a child to a parse tree
    '''node.</summary>
    '''
    '''<param name='node'>the parent node</param>
    '''<param name='child'>the child node, or null</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub ChildAdditiveExpression(ByVal node As Production, ByVal child As Node)
        node.AddChild(child)
    End Sub

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterAdditiveOp(ByVal node As Production)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitAdditiveOp(ByVal node As Production) As Node
        Return node
    End Function

    '''<summary>Called when adding a child to a parse tree
    '''node.</summary>
    '''
    '''<param name='node'>the parent node</param>
    '''<param name='child'>the child node, or null</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub ChildAdditiveOp(ByVal node As Production, ByVal child As Node)
        node.AddChild(child)
    End Sub

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterMultiplicativeExpression(ByVal node As Production)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitMultiplicativeExpression(ByVal node As Production) As Node
        Return node
    End Function

    '''<summary>Called when adding a child to a parse tree
    '''node.</summary>
    '''
    '''<param name='node'>the parent node</param>
    '''<param name='child'>the child node, or null</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub ChildMultiplicativeExpression(ByVal node As Production, ByVal child As Node)
        node.AddChild(child)
    End Sub

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterMultiplicativeOp(ByVal node As Production)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitMultiplicativeOp(ByVal node As Production) As Node
        Return node
    End Function

    '''<summary>Called when adding a child to a parse tree
    '''node.</summary>
    '''
    '''<param name='node'>the parent node</param>
    '''<param name='child'>the child node, or null</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub ChildMultiplicativeOp(ByVal node As Production, ByVal child As Node)
        node.AddChild(child)
    End Sub

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterExponentiationExpression(ByVal node As Production)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitExponentiationExpression(ByVal node As Production) As Node
        Return node
    End Function

    '''<summary>Called when adding a child to a parse tree
    '''node.</summary>
    '''
    '''<param name='node'>the parent node</param>
    '''<param name='child'>the child node, or null</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub ChildExponentiationExpression(ByVal node As Production, ByVal child As Node)
        node.AddChild(child)
    End Sub

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterPercentExpression(ByVal node As Production)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitPercentExpression(ByVal node As Production) As Node
        Return node
    End Function

    '''<summary>Called when adding a child to a parse tree
    '''node.</summary>
    '''
    '''<param name='node'>the parent node</param>
    '''<param name='child'>the child node, or null</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub ChildPercentExpression(ByVal node As Production, ByVal child As Node)
        node.AddChild(child)
    End Sub

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterUnaryExpression(ByVal node As Production)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitUnaryExpression(ByVal node As Production) As Node
        Return node
    End Function

    '''<summary>Called when adding a child to a parse tree
    '''node.</summary>
    '''
    '''<param name='node'>the parent node</param>
    '''<param name='child'>the child node, or null</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub ChildUnaryExpression(ByVal node As Production, ByVal child As Node)
        node.AddChild(child)
    End Sub

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterBasicExpression(ByVal node As Production)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitBasicExpression(ByVal node As Production) As Node
        Return node
    End Function

    '''<summary>Called when adding a child to a parse tree
    '''node.</summary>
    '''
    '''<param name='node'>the parent node</param>
    '''<param name='child'>the child node, or null</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub ChildBasicExpression(ByVal node As Production, ByVal child As Node)
        node.AddChild(child)
    End Sub

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterReference(ByVal node As Production)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitReference(ByVal node As Production) As Node
        Return node
    End Function

    '''<summary>Called when adding a child to a parse tree
    '''node.</summary>
    '''
    '''<param name='node'>the parent node</param>
    '''<param name='child'>the child node, or null</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub ChildReference(ByVal node As Production, ByVal child As Node)
        node.AddChild(child)
    End Sub

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterGridReferenceExpression(ByVal node As Production)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitGridReferenceExpression(ByVal node As Production) As Node
        Return node
    End Function

    '''<summary>Called when adding a child to a parse tree
    '''node.</summary>
    '''
    '''<param name='node'>the parent node</param>
    '''<param name='child'>the child node, or null</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub ChildGridReferenceExpression(ByVal node As Production, ByVal child As Node)
        node.AddChild(child)
    End Sub

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterGridReference(ByVal node As Production)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitGridReference(ByVal node As Production) As Node
        Return node
    End Function

    '''<summary>Called when adding a child to a parse tree
    '''node.</summary>
    '''
    '''<param name='node'>the parent node</param>
    '''<param name='child'>the child node, or null</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub ChildGridReference(ByVal node As Production, ByVal child As Node)
        node.AddChild(child)
    End Sub

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterFunctionCall(ByVal node As Production)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitFunctionCall(ByVal node As Production) As Node
        Return node
    End Function

    '''<summary>Called when adding a child to a parse tree
    '''node.</summary>
    '''
    '''<param name='node'>the parent node</param>
    '''<param name='child'>the child node, or null</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub ChildFunctionCall(ByVal node As Production, ByVal child As Node)
        node.AddChild(child)
    End Sub

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterArgumentList(ByVal node As Production)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitArgumentList(ByVal node As Production) As Node
        Return node
    End Function

    '''<summary>Called when adding a child to a parse tree
    '''node.</summary>
    '''
    '''<param name='node'>the parent node</param>
    '''<param name='child'>the child node, or null</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub ChildArgumentList(ByVal node As Production, ByVal child As Node)
        node.AddChild(child)
    End Sub

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterExpressionGroup(ByVal node As Production)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitExpressionGroup(ByVal node As Production) As Node
        Return node
    End Function

    '''<summary>Called when adding a child to a parse tree
    '''node.</summary>
    '''
    '''<param name='node'>the parent node</param>
    '''<param name='child'>the child node, or null</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub ChildExpressionGroup(ByVal node As Production, ByVal child As Node)
        node.AddChild(child)
    End Sub

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterPrimitive(ByVal node As Production)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitPrimitive(ByVal node As Production) As Node
        Return node
    End Function

    '''<summary>Called when adding a child to a parse tree
    '''node.</summary>
    '''
    '''<param name='node'>the parent node</param>
    '''<param name='child'>the child node, or null</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub ChildPrimitive(ByVal node As Production, ByVal child As Node)
        node.AddChild(child)
    End Sub

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterBoolean(ByVal node As Production)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitBoolean(ByVal node As Production) As Node
        Return node
    End Function

    '''<summary>Called when adding a child to a parse tree
    '''node.</summary>
    '''
    '''<param name='node'>the parent node</param>
    '''<param name='child'>the child node, or null</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub ChildBoolean(ByVal node As Production, ByVal child As Node)
        node.AddChild(child)
    End Sub

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub EnterErrorExpression(ByVal node As Production)
    End Sub

    '''<summary>Called when exiting a parse tree node.</summary>
    '''
    '''<param name='node'>the node being exited</param>
    '''
    '''<returns>the node to add to the parse tree, or
    '''         null if no parse tree should be created</returns>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Function ExitErrorExpression(ByVal node As Production) As Node
        Return node
    End Function

    '''<summary>Called when adding a child to a parse tree
    '''node.</summary>
    '''
    '''<param name='node'>the parent node</param>
    '''<param name='child'>the child node, or null</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overridable Sub ChildErrorExpression(ByVal node As Production, ByVal child As Node)
        node.AddChild(child)
    End Sub
End Class
