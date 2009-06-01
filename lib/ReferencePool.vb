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
' FormulaEngine - A library for parsing and managing formulas
' Copyright © 2007 Eugene Ciloci
'

Imports System.Collections.Generic

''' <summary>
''' Maintains a pool of all references.  This allows formulas to re-use instances of a reference so that when a reference changes, all formulas
''' that use it will see the change.  It uses simple reference counting to ensure that unused references are released.
''' </summary>
<Serializable()> _
Friend Class ReferencePool
	Implements System.Runtime.Serialization.ISerializable

	Private MyReferenceMap As IDictionary(Of Reference, ReferencePoolInfo)
	Private MyOwner As FormulaEngine
	Private Const VERSION As Integer = 1

	Public Sub New(ByVal owner As FormulaEngine)
		MyReferenceMap = New Dictionary(Of Reference, ReferencePoolInfo)(New ReferenceEqualityComparer())
		MyOwner = owner
	End Sub

	Private Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
		MyOwner = info.GetValue("Engine", GetType(FormulaEngine))
		MyReferenceMap = info.GetValue("Map", GetType(IDictionary(Of Reference, ReferencePoolInfo)))
	End Sub

	Public Sub GetObjectData(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext) Implements System.Runtime.Serialization.ISerializable.GetObjectData
		info.AddValue("Version", VERSION)
		info.AddValue("Engine", MyOwner)
		info.AddValue("Map", MyReferenceMap)
	End Sub

	Public Function GetReference(ByVal target As Reference) As Reference
		Debug.Assert(target.Valid = True, "expecting reference to be valid")
		Dim ref As Reference = Me.GetPooledReference(target)
		If ref Is Nothing Then
			Me.AddReference(target)
			Return target
		Else
			Me.IncrementReferenceCount(ref)
			Return ref
		End If
	End Function

	Private Sub IncrementReferenceCount(ByVal ref As Reference)
		Dim info As ReferencePoolInfo = MyReferenceMap.Item(ref)
		info.Count += 1
		MyReferenceMap.Item(ref) = info
	End Sub

	Private Sub AddReference(ByVal ref As Reference)
		Dim info As New ReferencePoolInfo(ref)
		MyReferenceMap.Add(ref, info)
		ref.SetEngine(MyOwner)
	End Sub

	Public Sub RemoveReference(ByVal ref As Reference)
		If ref.Valid = False Then
			Return
		ElseIf MyReferenceMap.ContainsKey(ref) = False Then
			Throw New ArgumentException("Reference is not in pool")
		End If

		Dim info As ReferencePoolInfo = MyReferenceMap.Item(ref)
		info.Count -= 1

		If info.Count = 0 Then
			MyReferenceMap.Remove(ref)
		Else
			MyReferenceMap.Item(ref) = info
		End If
	End Sub

	Public Function GetPooledReference(ByVal target As Reference) As Reference
		If MyReferenceMap.ContainsKey(target) = True Then
			Dim info As ReferencePoolInfo = MyReferenceMap.Item(target)
			Return info.Target
		Else
			Return Nothing
		End If
	End Function

	Public Sub Clear()
		MyReferenceMap.Clear()
	End Sub

	''' <summary>
	''' Perform an operation on a set of references
	''' </summary>
	Public Sub DoReferenceOperation(ByVal targets As IList, ByVal refOp As ReferenceOperator)
		' Keep track of references in various states
		Dim affectedValid As IList = New ArrayList
		Dim invalid As IList = New ArrayList
		Dim affected As IList = New ArrayList
		refOp.PreOperate(targets)

		For Each ref As Reference In targets
			Dim result As ReferenceOperationResultType = refOp.Operate(ref)

			If result <> ReferenceOperationResultType.NotAffected Then
				' Note references that are affected (valid or not)
				affected.Add(ref)
			End If

			If result = ReferenceOperationResultType.Invalidated Then
				' Note invalid references
				invalid.Add(ref)
			ElseIf result = ReferenceOperationResultType.Affected Then
				' Note affected but valid references
				affectedValid.Add(ref)
			End If
		Next

		Me.ProcessInvalidReferences(invalid)
		Me.ProcessAffectedValidReferences(affectedValid)
		MyOwner.RemoveInvalidFormulas(invalid)
		MyOwner.RecalculateAffectedReferences(affected)
		refOp.PostOperate(affectedValid)
	End Sub

	Private Sub ProcessInvalidReferences(ByVal invalid As IList)
		For Each ref As Reference In invalid
			MyReferenceMap.Remove(ref)
			ref.MarkAsInvalid()
		Next
	End Sub

	' Go through all references that were affected by an operation and rehash them
	Private Sub ProcessAffectedValidReferences(ByVal affectedValid As IList)
		For Each ref As Reference In affectedValid
			Me.RehashReference(ref)
		Next
	End Sub

	Private Sub RehashReference(ByVal ref As Reference)
		Dim info As ReferencePoolInfo = MyReferenceMap.Item(ref)
		MyReferenceMap.Remove(ref)
		ref.ComputeHashCode()
		MyReferenceMap.Add(ref, info)
	End Sub

    Public Function FindReferences(ByVal predicate As ReferencePredicateBase) As IList(Of Reference)
        Dim found As IList = New List(Of Reference)

        For Each ref As Reference In MyReferenceMap.Keys
            ' Ignore invalid references
            If ref.Valid = True AndAlso predicate.IsMatch(ref) = True Then
                found.Add(ref)
            End If
        Next

        Return found
    End Function

	Public ReadOnly Property ReferenceCount() As Integer
		Get
			Return MyReferenceMap.Count
		End Get
	End Property
End Class