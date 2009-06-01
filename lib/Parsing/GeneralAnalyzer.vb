' GeneralAnalyzer.vb
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
Friend MustInherit Class GeneralAnalyzer
    Inherits Analyzer

    '''<summary>Called when entering a parse tree node.</summary>
    '''
    '''<param name='node'>the node being entered</param>
    '''
    '''<exception cref='ParseException'>if the node analysis
    '''discovered errors</exception>
    Public Overrides Sub Enter(ByVal node As Node)
        Select Case node.Id
        Case GeneralConstants.ADD
            EnterAdd(CType(node,Token))

        Case GeneralConstants.SUB
            EnterSub(CType(node,Token))

        Case GeneralConstants.MUL
            EnterMul(CType(node,Token))

        Case GeneralConstants.DIV
            EnterDiv(CType(node,Token))

        Case GeneralConstants.EXP
            EnterExp(CType(node,Token))

        Case GeneralConstants.CONCAT
            EnterConcat(CType(node,Token))

        Case GeneralConstants.LEFT_PAREN
            EnterLeftParen(CType(node,Token))

        Case GeneralConstants.RIGHT_PAREN
            EnterRightParen(CType(node,Token))

        Case GeneralConstants.PERCENT
            EnterPercent(CType(node,Token))

        Case GeneralConstants.ARG_SEPARATOR
            EnterArgSeparator(CType(node,Token))

        Case GeneralConstants.EQ
            EnterEq(CType(node,Token))

        Case GeneralConstants.LT
            EnterLt(CType(node,Token))

        Case GeneralConstants.GT
            EnterGt(CType(node,Token))

        Case GeneralConstants.LTE
            EnterLte(CType(node,Token))

        Case GeneralConstants.GTE
            EnterGte(CType(node,Token))

        Case GeneralConstants.NE
            EnterNe(CType(node,Token))

        Case GeneralConstants.STRING_LITERAL
            EnterStringLiteral(CType(node,Token))

        Case GeneralConstants.NUMBER
            EnterNumber(CType(node,Token))

        Case GeneralConstants.TRUE
            EnterTrue(CType(node,Token))

        Case GeneralConstants.FALSE
            EnterFalse(CType(node,Token))

        Case GeneralConstants.FUNCTION_NAME
            EnterFunctionName(CType(node,Token))

        Case GeneralConstants.DEFINED_NAME
            EnterDefinedName(CType(node,Token))

        Case GeneralConstants.FORMULA
            EnterFormula(CType(node,Production))

        Case GeneralConstants.EXPRESSION
            EnterExpression(CType(node,Production))

        Case GeneralConstants.LOGICAL_EXPRESSION
            EnterLogicalExpression(CType(node,Production))

        Case GeneralConstants.LOGICAL_OP
            EnterLogicalOp(CType(node,Production))

        Case GeneralConstants.CONCAT_EXPRESSION
            EnterConcatExpression(CType(node,Production))

        Case GeneralConstants.ADDITIVE_EXPRESSION
            EnterAdditiveExpression(CType(node,Production))

        Case GeneralConstants.ADDITIVE_OP
            EnterAdditiveOp(CType(node,Production))

        Case GeneralConstants.MULTIPLICATIVE_EXPRESSION
            EnterMultiplicativeExpression(CType(node,Production))

        Case GeneralConstants.MULTIPLICATIVE_OP
            EnterMultiplicativeOp(CType(node,Production))

        Case GeneralConstants.EXPONENTIATION_EXPRESSION
            EnterExponentiationExpression(CType(node,Production))

        Case GeneralConstants.PERCENT_EXPRESSION
            EnterPercentExpression(CType(node,Production))

        Case GeneralConstants.UNARY_EXPRESSION
            EnterUnaryExpression(CType(node,Production))

        Case GeneralConstants.BASIC_EXPRESSION
            EnterBasicExpression(CType(node,Production))

        Case GeneralConstants.FUNCTION_CALL
            EnterFunctionCall(CType(node,Production))

        Case GeneralConstants.ARGUMENT_LIST
            EnterArgumentList(CType(node,Production))

        Case GeneralConstants.EXPRESSION_GROUP
            EnterExpressionGroup(CType(node,Production))

        Case GeneralConstants.PRIMITIVE
            EnterPrimitive(CType(node,Production))

        Case GeneralConstants.BOOLEAN
            EnterBoolean(CType(node,Production))

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
        Case GeneralConstants.ADD
            return ExitAdd(CType(node,Token))

        Case GeneralConstants.SUB
            return ExitSub(CType(node,Token))

        Case GeneralConstants.MUL
            return ExitMul(CType(node,Token))

        Case GeneralConstants.DIV
            return ExitDiv(CType(node,Token))

        Case GeneralConstants.EXP
            return ExitExp(CType(node,Token))

        Case GeneralConstants.CONCAT
            return ExitConcat(CType(node,Token))

        Case GeneralConstants.LEFT_PAREN
            return ExitLeftParen(CType(node,Token))

        Case GeneralConstants.RIGHT_PAREN
            return ExitRightParen(CType(node,Token))

        Case GeneralConstants.PERCENT
            return ExitPercent(CType(node,Token))

        Case GeneralConstants.ARG_SEPARATOR
            return ExitArgSeparator(CType(node,Token))

        Case GeneralConstants.EQ
            return ExitEq(CType(node,Token))

        Case GeneralConstants.LT
            return ExitLt(CType(node,Token))

        Case GeneralConstants.GT
            return ExitGt(CType(node,Token))

        Case GeneralConstants.LTE
            return ExitLte(CType(node,Token))

        Case GeneralConstants.GTE
            return ExitGte(CType(node,Token))

        Case GeneralConstants.NE
            return ExitNe(CType(node,Token))

        Case GeneralConstants.STRING_LITERAL
            return ExitStringLiteral(CType(node,Token))

        Case GeneralConstants.NUMBER
            return ExitNumber(CType(node,Token))

        Case GeneralConstants.TRUE
            return ExitTrue(CType(node,Token))

        Case GeneralConstants.FALSE
            return ExitFalse(CType(node,Token))

        Case GeneralConstants.FUNCTION_NAME
            return ExitFunctionName(CType(node,Token))

        Case GeneralConstants.DEFINED_NAME
            return ExitDefinedName(CType(node,Token))

        Case GeneralConstants.FORMULA
            return ExitFormula(CType(node,Production))

        Case GeneralConstants.EXPRESSION
            return ExitExpression(CType(node,Production))

        Case GeneralConstants.LOGICAL_EXPRESSION
            return ExitLogicalExpression(CType(node,Production))

        Case GeneralConstants.LOGICAL_OP
            return ExitLogicalOp(CType(node,Production))

        Case GeneralConstants.CONCAT_EXPRESSION
            return ExitConcatExpression(CType(node,Production))

        Case GeneralConstants.ADDITIVE_EXPRESSION
            return ExitAdditiveExpression(CType(node,Production))

        Case GeneralConstants.ADDITIVE_OP
            return ExitAdditiveOp(CType(node,Production))

        Case GeneralConstants.MULTIPLICATIVE_EXPRESSION
            return ExitMultiplicativeExpression(CType(node,Production))

        Case GeneralConstants.MULTIPLICATIVE_OP
            return ExitMultiplicativeOp(CType(node,Production))

        Case GeneralConstants.EXPONENTIATION_EXPRESSION
            return ExitExponentiationExpression(CType(node,Production))

        Case GeneralConstants.PERCENT_EXPRESSION
            return ExitPercentExpression(CType(node,Production))

        Case GeneralConstants.UNARY_EXPRESSION
            return ExitUnaryExpression(CType(node,Production))

        Case GeneralConstants.BASIC_EXPRESSION
            return ExitBasicExpression(CType(node,Production))

        Case GeneralConstants.FUNCTION_CALL
            return ExitFunctionCall(CType(node,Production))

        Case GeneralConstants.ARGUMENT_LIST
            return ExitArgumentList(CType(node,Production))

        Case GeneralConstants.EXPRESSION_GROUP
            return ExitExpressionGroup(CType(node,Production))

        Case GeneralConstants.PRIMITIVE
            return ExitPrimitive(CType(node,Production))

        Case GeneralConstants.BOOLEAN
            return ExitBoolean(CType(node,Production))

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
        Case GeneralConstants.FORMULA
            ChildFormula(node, child)

        Case GeneralConstants.EXPRESSION
            ChildExpression(node, child)

        Case GeneralConstants.LOGICAL_EXPRESSION
            ChildLogicalExpression(node, child)

        Case GeneralConstants.LOGICAL_OP
            ChildLogicalOp(node, child)

        Case GeneralConstants.CONCAT_EXPRESSION
            ChildConcatExpression(node, child)

        Case GeneralConstants.ADDITIVE_EXPRESSION
            ChildAdditiveExpression(node, child)

        Case GeneralConstants.ADDITIVE_OP
            ChildAdditiveOp(node, child)

        Case GeneralConstants.MULTIPLICATIVE_EXPRESSION
            ChildMultiplicativeExpression(node, child)

        Case GeneralConstants.MULTIPLICATIVE_OP
            ChildMultiplicativeOp(node, child)

        Case GeneralConstants.EXPONENTIATION_EXPRESSION
            ChildExponentiationExpression(node, child)

        Case GeneralConstants.PERCENT_EXPRESSION
            ChildPercentExpression(node, child)

        Case GeneralConstants.UNARY_EXPRESSION
            ChildUnaryExpression(node, child)

        Case GeneralConstants.BASIC_EXPRESSION
            ChildBasicExpression(node, child)

        Case GeneralConstants.FUNCTION_CALL
            ChildFunctionCall(node, child)

        Case GeneralConstants.ARGUMENT_LIST
            ChildArgumentList(node, child)

        Case GeneralConstants.EXPRESSION_GROUP
            ChildExpressionGroup(node, child)

        Case GeneralConstants.PRIMITIVE
            ChildPrimitive(node, child)

        Case GeneralConstants.BOOLEAN
            ChildBoolean(node, child)

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
End Class
