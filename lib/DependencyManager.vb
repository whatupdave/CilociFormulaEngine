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
''' Manages formula dependencies
''' </summary>
<Serializable()> _
Friend Class DependencyManager
    Implements System.Runtime.Serialization.ISerializable

    ''' <summary>
    ''' A data structure to store a reference and a list of its dependents
    ''' </summary>
    <Serializable()> _
    Private Class DependencyMap

        Private MyMap As Dictionary(Of Reference, IList(Of Reference))

        Public Sub New()
            MyMap = New Dictionary(Of Reference, IList(Of Reference))
        End Sub

        Public Sub AddDependency(ByVal tail As Reference, ByVal head As Reference)
            Dim dependents As IList(Of Reference) = Me.GetDependentsList(tail)
            'Debug.Assert(dependents.Contains(head) = False, "head already in list")
            dependents.Add(head)
        End Sub

        Public Sub RemoveDependency(ByVal tail As Reference, ByVal head As Reference)
            Dim list As IList(Of Reference) = MyMap.Item(tail)
            'Debug.Assert(list.Contains(head), "reference not in list")
            list.Remove(head)

            If list.Count = 0 Then
                MyMap.Remove(tail)
            End If
        End Sub

        Public Function GetRawDependents(ByVal tail As Reference) As IList(Of Reference)
            Return MyMap.Item(tail)
        End Function

        Public Sub RestoreRawDependents(ByVal tail As Reference, ByVal dependents As IList(Of Reference))
            MyMap.Add(tail, dependents)
        End Sub

        Private Function GetDependentsList(ByVal tail As Reference) As IList(Of Reference)
            Dim list As IList(Of Reference)
            If Not MyMap.ContainsKey(tail) Then
                list = New List(Of Reference)
                MyMap.Add(tail, list)
            End If

            Return MyMap.Item(tail)
        End Function

        Public Sub RemoveTail(ByVal tail As Reference)
            Debug.Assert(MyMap.ContainsKey(tail), "tail not in map")
            MyMap.Remove(tail)
        End Sub

        Public Function ContainsTail(ByVal tail As Reference) As Boolean
            Return MyMap.ContainsKey(tail)
        End Function

        Public Sub Clear()
            MyMap.Clear()
        End Sub

        Public Function IsCircularReference(ByVal tail As Reference) As Boolean
            Return Me.IsCircularReferenceInternal(tail, tail)
        End Function

        Private Function IsCircularReferenceInternal(ByVal target As Reference, ByVal currentReference As Object) As Boolean
            If Not MyMap.ContainsKey(currentReference) Then
                Return False
            End If

            Dim directDependents As IList(Of Reference) = MyMap.Item(currentReference)

            For Each dependant As Reference In directDependents
                If target.IsReferenceEqualForCircularReference(dependant) = True Then
                    Return True
                ElseIf Me.IsCircularReferenceInternal(target, dependant) = True Then
                    Return True
                End If
            Next

            Return False
        End Function

        Public Sub GetDirectDependents(ByVal tail As Reference, ByVal dest As IList(Of Reference))
            If Not MyMap.ContainsKey(tail) Then
                Return
            End If
            Dim list As IList(Of Reference) = MyMap.Item(tail)
            For Each o As Object In list
                dest.Add(o)
            Next
        End Sub

        ''' <summary>
        ''' Copy a portion of our dependencies into another DependencyMap
        ''' </summary>
        Public Sub CloneSourceDependents(ByVal sources As IList(Of Reference), ByVal dependents As DependencyMap, ByVal precedents As DependencyMap)
            Dim seenNodes As IDictionary = New Hashtable
            For Each source As Reference In sources
                Me.CloneDependentsInternal(source, dependents, precedents, seenNodes)
            Next
        End Sub

        Private Sub CloneDependentsInternal(ByVal tail As Reference, ByVal dependents As DependencyMap, ByVal precedents As DependencyMap, ByVal seenNodes As IDictionary)
            ' Have we seen this reference already?
            If seenNodes.Contains(tail) = True Then
                ' Yes so just return
                Return
            Else
                ' Mark it as seen
                seenNodes.Add(tail, Nothing)
            End If

            If Not MyMap.ContainsKey(tail) Then
                Return
            End If

            Dim list As IList(Of Reference) = MyMap.Item(tail)

            For Each dependant As Reference In list
                dependents.AddDependency(tail, dependant)
                precedents.AddDependency(dependant, tail)
                Me.CloneDependentsInternal(dependant, dependents, precedents, seenNodes)
            Next
        End Sub

        Public Overrides Function ToString() As String
            Dim lines(MyMap.Count - 1) As String
            Dim index As Integer

            For Each de As KeyValuePair(Of Reference, IList(Of Reference)) In MyMap
                Dim key As Reference = de.Key
                Dim list As IList(Of Reference) = de.Value
                lines(index) = String.Format("{0} -> {1}", key, Reference.References2String(list))
                index += 1
            Next

            Return String.Join(System.Environment.NewLine, lines)
        End Function

        Public ReadOnly Property DependencyCount() As Integer
            Get
                Dim count As Integer

                For Each l As IList(Of Reference) In MyMap.Values
                    count += l.Count
                Next

                Return count
            End Get
        End Property
    End Class

    Private MustInherit Class DependencyReferencePredicateBase
        Inherits ReferencePredicateBase

        Private MyOwner As DependencyManager

        Public Sub New()

        End Sub

        Public Sub SetOwner(ByVal owner As DependencyManager)
            MyOwner = owner
        End Sub

        Protected ReadOnly Property Owner() As DependencyManager
            Get
                Return MyOwner
            End Get
        End Property
    End Class

    Private Class IntersectingSourcePredicate
        Inherits DependencyReferencePredicateBase

        Private MyTarget As Reference

        Public Sub New(ByVal target As Reference)
            MyTarget = target
        End Sub

        Public Overrides Function IsMatch(ByVal ref As Reference) As Boolean
            Return MyTarget.Intersects(ref) And Me.Owner.IsSource(ref)
        End Function
    End Class

    Private Class RangeLinkPredicate
        Inherits DependencyReferencePredicateBase

        Public Overrides Function IsMatch(ByVal ref As Reference) As Boolean
            Return ref.CanRangeLink
        End Function
    End Class

    Private Class LinkerIntersectPredicate
        Inherits DependencyReferencePredicateBase

        Private MyTarget As Reference

        Public Sub New(ByVal target As Reference)
            MyTarget = target
        End Sub

        Public Overrides Function IsMatch(ByVal ref As Reference) As Boolean
            Return ref.CanRangeLink = False AndAlso MyTarget.Intersects(ref) AndAlso Me.Owner.IsSource(ref) = False
        End Function
    End Class

    Private Class SourcePredicate
        Inherits DependencyReferencePredicateBase

        Public Sub New()

        End Sub

        Public Overrides Function IsMatch(ByVal ref As Reference) As Boolean
            Return Owner.IsSource(ref)
        End Function
    End Class

    Private Class IntersectingPredicate
        Inherits DependencyReferencePredicateBase

        Private MyTarget As Reference

        Public Sub New(ByVal target As Reference)
            MyTarget = target
        End Sub

        Public Overrides Function IsMatch(ByVal ref As Reference) As Boolean
            Return MyTarget.Intersects(ref)
        End Function
    End Class

    Private Delegate Sub IntersectingDependantProcessor(ByVal tail As Reference, ByVal dependant As Reference)

    Private MyDependents As DependencyMap
    Private MyPrecedents As DependencyMap
    Private MyOwner As FormulaEngine
    Private MySuspendRangeLinks As Boolean
    Private Const VERSION As Integer = 1

    Public Sub New(ByVal owner As FormulaEngine)
        MyOwner = owner
        MyDependents = New DependencyMap
        MyPrecedents = New DependencyMap
    End Sub

    Private Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
        MyDependents = info.GetValue("Dependents", GetType(DependencyMap))
        MyPrecedents = info.GetValue("Precedents", GetType(DependencyMap))
        MyOwner = info.GetValue("Engine", GetType(FormulaEngine))
    End Sub

    Public Overridable Sub GetObjectData(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext) Implements System.Runtime.Serialization.ISerializable.GetObjectData
        info.AddValue("Version", VERSION)
        info.AddValue("Dependents", MyDependents)
        info.AddValue("Precedents", MyPrecedents)
        info.AddValue("Engine", MyOwner)
    End Sub

    ''' <summary>
    ''' Add all dependencies of a formula
    ''' </summary>
    Public Sub AddFormula(ByVal f As Formula)
        Me.RemoveRangeLinks()

        ' Add dependencies from the formula's self reference to each of its references
        Dim selfRef As Reference = f.SelfReferenceInternal
        Dim refs As Reference() = f.DependencyReferences

        For i As Integer = 0 To refs.Length - 1
            Dim ref As Reference = refs(i)
            MyDependents.AddDependency(ref, selfRef)
            MyPrecedents.AddDependency(selfRef, ref)
        Next

        Me.AddRangeLinks()
    End Sub

    ''' <summary>
    ''' Remove all dependencies of a formula
    ''' </summary>
    Public Sub RemoveFormula(ByVal f As Formula)
        Dim selfRef As Reference = f.SelfReferenceInternal
        Me.RemoveRangeLinks()

        For Each ref As Reference In f.DependencyReferences
            Me.RemoveDependency(ref, selfRef, MyDependents, MyPrecedents)
        Next

        Me.AddRangeLinks()
    End Sub

    Public Function FormulaHasCircularReference(ByVal f As Formula) As Boolean
        Return Me.IsCircularReference(f.SelfReferenceInternal)
    End Function

    Public Function IsCircularReference(ByVal ref As Reference) As Boolean
        Return MyDependents.IsCircularReference(ref)
    End Function

    Public Sub AddRangeLinks()
        If MySuspendRangeLinks = False Then
            'Me.ProcessRangeLinks(AddressOf DoAddIntersectingDependency)
        End If
    End Sub

    Public Sub RemoveRangeLinks()
        If MySuspendRangeLinks = False Then
            'Me.ProcessRangeLinks(AddressOf DoRemoveIntersectingDependency)
        End If
    End Sub

    ''' <summary>
    ''' Creates or removes range links
    ''' </summary>
    ''' <remarks>
    ''' A range link is an optimization for dependencies involving ranges.  Given a reference to a range, we find
    ''' all non-range references that intersect it.  We then create dependencies (or links) between those references and
    ''' all depedendents of the target range.  If we didn't do this, then when walking a dependency chain, we'd have
    ''' to find all intersecting references for each reference along the chain.
    ''' </remarks>
    Private Sub ProcessRangeLinks(ByVal processor As IntersectingDependantProcessor)
        Dim linkers As IList(Of Reference) = Me.FindReferences(New RangeLinkPredicate)

        For Each linker As Reference In linkers
            Dim intersecting As IList(Of Reference) = Me.FindReferences(New LinkerIntersectPredicate(linker))
            Me.ProcessLinkerIntersects(linker, intersecting, processor)
        Next
    End Sub

    Private Sub ProcessLinkerIntersects(ByVal linker As Reference, ByVal intersecting As IList(Of Reference), ByVal processor As IntersectingDependantProcessor)
        Dim linkerDependents As IList(Of Reference) = New List(Of Reference)
        MyDependents.GetDirectDependents(linker, linkerDependents)

        For Each intersect As Reference In intersecting
            Me.ProcessLinkerDependents(intersect, linkerDependents, processor)
        Next
    End Sub

    Private Sub ProcessLinkerDependents(ByVal intersect As Reference, ByVal linkerDependents As IList(Of Reference), ByVal processor As IntersectingDependantProcessor)
        For Each linkerDependant As Reference In linkerDependents
            processor(intersect, linkerDependant)
        Next
    End Sub

    Private Sub DoAddIntersectingDependency(ByVal tail As Reference, ByVal dependant As Reference)
        MyDependents.AddDependency(tail, dependant)
        MyPrecedents.AddDependency(dependant, tail)
    End Sub

    Private Sub DoRemoveIntersectingDependency(ByVal tail As Reference, ByVal dependant As Reference)
        MyDependents.RemoveDependency(tail, dependant)
        MyPrecedents.RemoveDependency(dependant, tail)
    End Sub

    Private Function FindReferences(ByVal predicate As DependencyReferencePredicateBase) As IList(Of Reference)
        predicate.SetOwner(Me)
        Return MyOwner.ReferencePool.FindReferences(predicate)
    End Function

    Private Sub RemoveDependency(ByVal tail As Reference, ByVal head As Reference, ByVal map As DependencyMap, ByVal inverseMap As DependencyMap)
        map.RemoveDependency(tail, head)
        inverseMap.RemoveDependency(head, tail)
    End Sub

    Friend Sub Clear()
        MyDependents.Clear()
        MyPrecedents.Clear()
    End Sub

    Public Function GetSources(ByVal refs As IList(Of Reference)) As IList(Of Reference)
        Dim sources As IList(Of Reference) = New List(Of Reference)

        For Each ref As Reference In refs
            If Me.IsSource(ref) = True Then
                sources.Add(ref)
            End If
        Next

        Return sources
    End Function

    Public Function GetReferenceCalculationList(ByVal root As Reference) As Reference()
        Dim roots As IList(Of Reference) = Me.FindReferences(New IntersectingPredicate(root))
        Return Me.GetCalculationList(roots)
    End Function

    Public Function GetAllCalculationList() As Reference()
        Dim roots As IList(Of Reference) = Me.FindReferences(New SourcePredicate)
        Return Me.GetCalculationList(roots)
    End Function

    ''' <summary>
    ''' Gets a sorted list of all references that depend on a given set of references
    ''' </summary>
    Public Function GetCalculationList(ByVal roots As IList(Of Reference)) As Reference()
        ' We need to get rid of all circular references from the roots list
        Me.RemoveCircularReferences(roots)
        ' Add any volatile functions to the list
        Me.AddVolatileReference(roots)
        ' Create temporary maps for calculating the list
        Dim tempDependents As New DependencyMap
        Dim tempPrecedents As New DependencyMap

        ' Clone the portion of our dependency map that contains the roots
        MyDependents.CloneSourceDependents(roots, tempDependents, tempPrecedents)

        ' We don't care about non-sources
        Me.RemoveNonSourcesFromRoots(roots, tempPrecedents)

        ' Sort the dependencies
        Dim targets As IList(Of Reference) = Me.TopologicalSort(tempDependents, tempPrecedents, roots)
        ' Sources can't be recalculated since they have no formula so get rid of them
        Me.RemoveSources(roots, targets)

        Dim arr(targets.Count - 1) As Reference
        targets.CopyTo(arr, 0)
        Return arr
    End Function

    Private Sub RemoveCircularReferences(ByVal roots As IList(Of Reference))
        Dim arr(roots.Count - 1) As Reference
        roots.CopyTo(arr, 0)

        For Each ref As Reference In arr
            If Me.IsCircularReference(ref) = True Then
                roots.Remove(ref)
            End If
        Next
    End Sub

    Private Sub RemoveNonSourcesFromRoots(ByVal roots As IList(Of Reference), ByVal precedents As DependencyMap)
        Dim arr(roots.Count - 1) As Reference
        roots.CopyTo(arr, 0)

        For i As Integer = 0 To arr.Length - 1
            Dim root As Reference = arr(i)
            If precedents.ContainsTail(root) = True Then
                roots.Remove(root)
            End If
        Next
    End Sub

    Private Sub RemoveSources(ByVal sources As IList(Of Reference), ByVal targets As IList(Of Reference))
        For Each source As Reference In sources
            targets.Remove(source)
        Next
    End Sub

    ''' <summary>
    ''' Add any volatile functions to the roots of a calculation list
    ''' </summary>
    Private Sub AddVolatileReference(ByVal roots As IList(Of Reference))
        Dim volatile As New VolatileFunctionReference()
        ' Are there any formulas that have volatile functions?
        volatile = MyOwner.ReferencePool.GetPooledReference(volatile)
        If Not volatile Is Nothing Then
            roots.Add(volatile)
        End If
    End Sub

    ''' <summary>
    ''' Perform the topological sort algorithm on our graph
    ''' </summary>
    Private Function TopologicalSort(ByVal dependents As DependencyMap, ByVal precedents As DependencyMap, ByVal sources As IList(Of Reference)) As IList(Of Reference)
        Dim q As New Queue(sources)
        Dim output As IList(Of Reference) = New List(Of Reference)
        Dim directDependents As IList(Of Reference) = New List(Of Reference)

        While q.Count > 0
            Dim n As Reference = q.Dequeue()
            output.Add(n)

            directDependents.Clear()
            dependents.GetDirectDependents(n, directDependents)

            For Each m As Reference In directDependents
                dependents.RemoveDependency(n, m)
                precedents.RemoveDependency(m, n)
                If precedents.ContainsTail(m) = False Then
                    q.Enqueue(m)
                End If
            Next
        End While

        ' We assume no circular references
        Return output
    End Function

    ''' <summary>
    ''' Determines if a reference is a source.  A source is a reference that only has dependents and no precedents
    ''' </summary>
    Private Function IsSource(ByVal ref As Reference) As Boolean
        Return MyPrecedents.ContainsTail(ref) = False
    End Function

    Public Function GetDirectDependentsCount(ByVal ref As Reference) As Integer
        Return Me.GetDirectDependentsInternal(ref, MyDependents)
    End Function

    Public Function GetDirectPrecedentsCount(ByVal ref As Reference) As Integer
        Return Me.GetDirectDependentsInternal(ref, MyPrecedents)
    End Function

    Private Function GetDirectDependentsInternal(ByVal ref As Reference, ByVal map As DependencyMap) As Integer
        Dim dependents As IList(Of Reference) = New List(Of Reference)
        map.GetDirectDependents(ref, dependents)
        Return dependents.Count
    End Function

    Public Sub SetSuspendRangeLinks(ByVal suspend As Boolean)
        MySuspendRangeLinks = suspend
    End Sub

    Public ReadOnly Property DependentsCount() As Integer
        Get
            Return MyDependents.DependencyCount
        End Get
    End Property

    Public ReadOnly Property PrecedentsCount() As Integer
        Get
            Return MyPrecedents.DependencyCount
        End Get
    End Property

    Public ReadOnly Property DependencyDump() As String
        Get
            Return MyDependents.ToString()
        End Get
    End Property
End Class