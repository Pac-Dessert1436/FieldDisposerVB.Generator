Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic
Imports IGIC = Microsoft.CodeAnalysis.IncrementalGeneratorInitializationContext
Imports System.Collections.Immutable

<Generator(LanguageNames.VisualBasic)>
Public NotInheritable Class DisposerGen
    Implements IIncrementalGenerator

    Private Enum TypeConstruct
        [Module] = 0
        [Class] = 1
        [Structure] = 2
    End Enum

    Private Class TypeInfo
        Public Property ContainingNamespace As String
        Public Property TypeName As String
        Public Property TypeConstruct As TypeConstruct
        Public Property Fields As List(Of FieldInfo)
        Public Property AccessModifier As String
        Public Property IsNested As Boolean
        Public Property ContainingType As TypeInfo
        Public Property NestedTypes As List(Of TypeInfo)
        Public Property HasDisposableBaseClass As Boolean
        Public Property IsSealedClass As Boolean
    End Class

    Private Class FieldInfo
        Public Property FieldName As String
        Public Property FieldType As String
        Public Property IsDisposable As Boolean
        Public Property IsNullable As Boolean
    End Class

    Public Sub Initialize(context As IGIC) Implements IIncrementalGenerator.Initialize
        ' Register the attribute source
        context.RegisterPostInitializationOutput(
            Sub(ctx) ctx.AddSource("DisposeFieldAttribute.g.vb", GenerateAttributeCode())
        )

        ' Find all partial types that have fields with DisposeField attribute
        Dim typeProvider = context.SyntaxProvider.CreateSyntaxProvider(
            Function(node, cancellationToken)
                Return TypeOf node Is Syntax.ClassStatementSyntax OrElse
                       TypeOf node Is Syntax.StructureStatementSyntax OrElse
                       TypeOf node Is Syntax.ModuleStatementSyntax
            End Function,
            Function(ctx, cancellationToken)
                Dim typeDecl = TryCast(ctx.Node, Syntax.TypeStatementSyntax)
                If typeDecl Is Nothing Then Return Nothing

                ' Get the containing type block (handles both top-level and nested types)
                Dim typeBlock = TryCast(typeDecl.Parent, Syntax.TypeBlockSyntax)
                If typeBlock Is Nothing Then Return Nothing

                ' Check for DisposeField attribute on any field
                For Each fieldDecl In typeBlock.Members.OfType(Of Syntax.FieldDeclarationSyntax)()
                    For Each attr In fieldDecl.AttributeLists
                        For Each attribute In attr.Attributes
                            Dim attrName = attribute.Name.ToString()
                            If attrName.Equals("DisposeField") OrElse attrName.EndsWith(".DisposeField") Then
                                Return typeBlock
                            End If
                        Next
                    Next
                Next

                Return Nothing
            End Function
        ).Where(Function(t) t IsNot Nothing)

        Dim compilation = context.CompilationProvider.Combine(typeProvider.Collect())
        context.RegisterSourceOutput(
            compilation, Sub(spc, tup) GenerateSource(tup.Left, tup.Right, spc))
    End Sub

    Private Function GenerateAttributeCode() As String
        Dim code As New System.Text.StringBuilder
        ' Add the attribute code
        code.AppendLine("Imports System

<AttributeUsage(AttributeTargets.Field, AllowMultiple:=False)>
Public NotInheritable Class DisposeFieldAttribute
    Inherits Attribute

    Public Sub New()
    End Sub
End Class" & vbCrLf)
        Dim ParamPatterns = Iterator Function()
                                ' Generate patterns for 0 to 5 parameters
                                For count = 0 To 5
                                    If count = 1 Then
                                        Yield ({"obj"}, {"T"})
                                        Continue For
                                    End If
                                    Dim params As New List(Of String)
                                    Dim generics As New List(Of String)

                                    For i = 1 To count
                                        params.Add($"arg{i}")
                                        generics.Add($"T{i}")
                                    Next

                                    Yield (params.ToArray(), generics.ToArray())
                                Next
                            End Function

        code.AppendLine("#Region ""Bonus: Disposable event classes""")

        ' Generate disposable event classes for different signatures
        For Each pattern In ParamPatterns()
            Dim paramGenerics = String.Format("(Of {0})", String.Join(", ", pattern.Item2))
            Dim paramList = String.Join(", ", pattern.Item1.Zip(pattern.Item2, Function(a, b) $"{a} As {b}"))
            Dim argumentList = String.Join(", ", pattern.Item1)
            If pattern.Item1.Count = 0 Then
                paramList = String.Empty
                paramGenerics = String.Empty
                argumentList = String.Empty
            End If
            code.AppendLine($"Public NotInheritable Class DisposableEvent{paramGenerics}")
            code.AppendLine("    Implements IDisposable")
            code.AppendLine()
            code.AppendLine($"    Private Event _SourceEvent As Action{paramGenerics}")
            code.AppendLine($"    Private _removalStack As New Stack(Of Action)")
            code.AppendLine("    Private _isDisposed As Boolean")
            code.AppendLine()
            code.AppendLine($"    Public Sub Subscribe(eventAction As Action{paramGenerics})")
            code.AppendLine($"        AddHandler _SourceEvent, eventAction")
            code.AppendLine("        _removalStack.Push(Sub() RemoveHandler _SourceEvent, eventAction)")
            code.AppendLine("    End Sub")
            code.AppendLine()
            code.AppendLine($"    Public Sub Publish({paramList})")
            code.AppendLine("        If _isDisposed Then Exit Sub")
            code.AppendLine($"        RaiseEvent _SourceEvent({argumentList})")
            code.AppendLine("    End Sub")
            code.AppendLine()
            code.AppendLine($"    Public ReadOnly Property EventCount As Integer")
            code.AppendLine("        Get")
            code.AppendLine($"            Return _removalStack.Count")
            code.AppendLine("        End Get")
            code.AppendLine("    End Property")
            code.AppendLine()
            code.AppendLine("    Public Sub Dispose() Implements IDisposable.Dispose")
            code.AppendLine("        Dispose(disposing:=True)")
            code.AppendLine("        GC.SuppressFinalize(Me)")
            code.AppendLine("    End Sub")
            code.AppendLine()
            code.AppendLine("    Private Sub Dispose(disposing As Boolean)")
            code.AppendLine("        If Not _isDisposed Then")
            code.AppendLine("            If disposing Then")
            code.AppendLine("                Do Until _removalStack.Count = 0")
            code.AppendLine("                    _removalStack.Pop().Invoke()")
            code.AppendLine("                Loop")
            code.AppendLine("            End If")
            code.AppendLine("            _isDisposed = True")
            code.AppendLine("        End If")
            code.AppendLine("    End Sub")
            code.AppendLine()
            code.AppendLine("    Protected Overrides Sub Finalize()")
            code.AppendLine("        Try")
            code.AppendLine("            Dispose(disposing:=False)")
            code.AppendLine("        Finally")
            code.AppendLine("            MyBase.Finalize()")
            code.AppendLine("        End Try")
            code.AppendLine("    End Sub")
            code.AppendLine("End Class")
            code.AppendLine()
        Next

        code.AppendLine("#End Region")
        Return code.ToString()
    End Function

    Private Sub GenerateSource(compilation As Compilation, typeBlocks As ImmutableArray(Of Syntax.TypeBlockSyntax), source As SourceProductionContext)
        ' Track processed files to avoid duplicates
        Dim processedFiles As New HashSet(Of String)()
        
        For Each typeBlock In typeBlocks
            If typeBlock IsNot Nothing Then
                Dim typeInfo = AnalyzeType(typeBlock, compilation, source.CancellationToken)
                If typeInfo IsNot Nothing AndAlso typeInfo.Fields.Count > 0 Then
                    ' Generate source code with complete type hierarchy for nested types
                    Dim sourceCode = GenerateDisposeCodeWithCompleteHierarchy(typeBlock, compilation, source.CancellationToken)
                    
                    ' Generate unique file name
                    Dim fileName As String = typeInfo.TypeName
                    If typeInfo.IsNested Then
                        ' For nested types, include the full hierarchy in the file name
                        Dim model = compilation.GetSemanticModel(typeBlock.SyntaxTree)
                        Dim typeSymbol = model.GetDeclaredSymbol(typeBlock.BlockStatement, source.CancellationToken)
                        If typeSymbol IsNot Nothing Then
                            fileName = typeSymbol.ToDisplayString().Replace(".", "_").Replace("+", "_")
                        End If
                    End If
                    
                    ' Skip if we've already generated this file
                    If processedFiles.Contains(fileName) Then Continue For
                    processedFiles.Add(fileName)
                    
                    source.AddSource($"{fileName}_Disposers.g.vb", sourceCode)
                End If
            End If
        Next
    End Sub

    Private Function AnalyzeType(typeBlock As Syntax.TypeBlockSyntax, compilation As Compilation, cancellationToken As Threading.CancellationToken) As TypeInfo
        Try
            Dim model = compilation.GetSemanticModel(typeBlock.SyntaxTree)
            Dim typeSymbol = model.GetDeclaredSymbol(typeBlock.BlockStatement, cancellationToken)
            If typeSymbol Is Nothing Then Return Nothing

            ' Get the containing namespace from the type block
            Dim containingNamespace As String = String.Empty
            Dim namespaceDecl = typeBlock.FirstAncestorOrSelf(Of Syntax.NamespaceBlockSyntax)()
            If namespaceDecl IsNot Nothing Then
                containingNamespace = namespaceDecl.NamespaceStatement.Name.ToString()
            End If

            Dim typeInfo As New TypeInfo With {
                .ContainingNamespace = containingNamespace,
                .TypeName = typeBlock.BlockStatement.Identifier.ValueText,
                .TypeConstruct = GetTypeConstruct(typeBlock),
                .AccessModifier = GetAccessModifier(typeSymbol),
                .IsNested = typeSymbol.ContainingType IsNot Nothing,
                .ContainingType = Nothing, ' This will be populated if we need to track hierarchy
                .NestedTypes = New List(Of TypeInfo)(),
                .HasDisposableBaseClass = CheckDisposableBaseClass(typeSymbol),
                .IsSealedClass = typeSymbol IsNot Nothing AndAlso typeSymbol.IsSealed,
                .Fields = New List(Of FieldInfo)()
            }

            For Each fieldDecl In typeBlock.Members.OfType(Of Syntax.FieldDeclarationSyntax)()
                For Each variable In fieldDecl.Declarators
                    Dim fieldName = variable.Names(0).Identifier.ValueText
                    Dim hasDisposeFieldAttribute = False
                    Dim fieldSymbol As IFieldSymbol = Nothing

                    ' Get the field symbol
                    For Each declarator In fieldDecl.Declarators
                        For Each name In declarator.Names
                            If name.Identifier.ValueText = fieldName Then
                                fieldSymbol = TryCast(model.GetDeclaredSymbol(name, cancellationToken), IFieldSymbol)
                                Exit For
                            End If
                        Next
                    Next

                    ' Check for DisposeField attribute
                    For Each attr In fieldDecl.AttributeLists
                        For Each attribute In attr.Attributes
                            Dim attrName = attribute.Name.ToString()
                            If attrName.Equals("DisposeField") OrElse attrName.EndsWith(".DisposeField") Then
                                hasDisposeFieldAttribute = True
                                Exit For
                            End If
                        Next
                        If hasDisposeFieldAttribute Then Exit For
                    Next

                    ' Check if the field type implements IDisposable
                    Dim isDisposable = False
                    Dim isNullable = False
                    If hasDisposeFieldAttribute AndAlso fieldSymbol IsNot Nothing Then
                        isDisposable = ImplementsIDisposable(fieldSymbol.Type)
                        ' Check if the field is nullable
                        isNullable = IsFieldNullable(fieldSymbol)
                    End If

                    If hasDisposeFieldAttribute Then
                        typeInfo.Fields.Add(New FieldInfo With {
                            .FieldName = fieldName,
                            .FieldType = If(fieldSymbol?.Type?.ToDisplayString(), "Object"),
                            .IsDisposable = isDisposable,
                            .IsNullable = isNullable
                        })
                    End If
                Next
            Next

            Return If(typeInfo.Fields.Count > 0, typeInfo, Nothing)
        Catch ex As Exception
            ' Log or ignore
            Return Nothing
        End Try
    End Function

    Private Function ImplementsIDisposable(typeSymbol As ITypeSymbol) As Boolean
        If typeSymbol Is Nothing Then Return False

        ' Check if the type itself is IDisposable
        If typeSymbol.ToDisplayString() = "System.IDisposable" Then Return True

        ' Check if it implements IDisposable
        For Each iface In typeSymbol.AllInterfaces
            If iface.ToDisplayString() = "System.IDisposable" Then
                Return True
            End If
        Next

        Return False
    End Function

    Private Function IsFieldNullable(fieldSymbol As IFieldSymbol) As Boolean
        If fieldSymbol Is Nothing Then Return False

        ' Check whether the field:
        ' - belongs to a nullable value type (Nullable(Of T))
        ' - has a type name contains "?" (VB.NET nullable syntax)
        ' - belongs to a reference type in itself (nullable by default in VB.NET)
        With fieldSymbol.Type
            If .OriginalDefinition.SpecialType = SpecialType.System_Nullable_T OrElse
                .ToDisplayString().Contains("?") OrElse .IsReferenceType Then
                Return True
            End If
        End With

        Return False
    End Function

    Private Function GetTypeConstruct(typeBlock As Syntax.TypeBlockSyntax) As TypeConstruct
        Select Case typeBlock.BlockStatement.Kind()
            Case SyntaxKind.ClassStatement
                Return TypeConstruct.Class
            Case SyntaxKind.StructureStatement
                Return TypeConstruct.Structure
            Case SyntaxKind.ModuleStatement
                Return TypeConstruct.Module
            Case Else
                Return TypeConstruct.Class
        End Select
    End Function

    Private Function CheckDisposableBaseClass(typeSymbol As INamedTypeSymbol) As Boolean
        If typeSymbol Is Nothing OrElse typeSymbol.BaseType Is Nothing Then Return False

        ' Check if the base class implements IDisposable
        Return ImplementsIDisposable(typeSymbol.BaseType)
    End Function

    Private Function GetAccessModifier(typeSymbol As INamedTypeSymbol) As String
        If typeSymbol Is Nothing Then Return String.Empty

        ' Check the declared accessibility
        Select Case typeSymbol.DeclaredAccessibility
            Case Accessibility.Public
                Return "Public"
            Case Accessibility.Private
                Return "Private"
            Case Accessibility.Internal
                Return "Friend"
            Case Else
                Return String.Empty
        End Select
    End Function

    Private Function GenerateDisposeCode(typeInfo As TypeInfo) As String
        Dim code As New System.Text.StringBuilder

        ' Add namespace if present
        If Not String.IsNullOrEmpty(typeInfo.ContainingNamespace) Then
            code.AppendLine($"Namespace {typeInfo.ContainingNamespace}")
            code.AppendLine()
        End If

        ' For nested classes, we need to generate the complete type hierarchy
        If typeInfo.IsNested Then
            ' Generate the complete type hierarchy for nested classes
            GenerateCompleteTypeHierarchy(code, typeInfo)
        Else
            ' For top-level types, just generate the type
            GenerateTypeHierarchy(code, typeInfo, 0)
        End If

        If Not String.IsNullOrEmpty(typeInfo.ContainingNamespace) Then
            code.AppendLine()
            code.AppendLine("End Namespace")
        End If

        Return code.ToString()
    End Function

    Private Function GenerateDisposeCodeWithCompleteHierarchy(typeBlock As Syntax.TypeBlockSyntax, compilation As Compilation, cancellationToken As Threading.CancellationToken) As String
        Dim code As New System.Text.StringBuilder
        
        ' Analyze the target type
        Dim typeInfo = AnalyzeType(typeBlock, compilation, cancellationToken)
        If typeInfo Is Nothing Then Return String.Empty

        ' Add namespace if present
        If Not String.IsNullOrEmpty(typeInfo.ContainingNamespace) Then
            code.AppendLine($"Namespace {typeInfo.ContainingNamespace}")
            code.AppendLine()
        End If

        ' For nested types, we need to generate the complete type hierarchy
        If typeInfo.IsNested Then
            ' Build the complete type hierarchy from the syntax tree
            GenerateCompleteTypeHierarchyFromSyntax(code, typeBlock, compilation, cancellationToken, typeInfo)
        Else
            ' For top-level types, just generate the type
            GenerateTypeHierarchy(code, typeInfo, 0)
        End If

        If Not String.IsNullOrEmpty(typeInfo.ContainingNamespace) Then
            code.AppendLine()
            code.AppendLine("End Namespace")
        End If

        Return code.ToString()
    End Function

    Private Sub GenerateCompleteTypeHierarchy(code As System.Text.StringBuilder, typeInfo As TypeInfo)
        ' For nested classes, we need to generate the complete type hierarchy
        ' This is a simplified approach - in a real implementation, we'd need to properly
        ' track the parent type hierarchy from the syntax tree
        
        ' For now, we'll generate just the nested type itself
        ' The parent types should already exist in the source code
        GenerateTypeHierarchy(code, typeInfo, 0)
    End Sub

    Private Sub GenerateCompleteTypeHierarchyFromSyntax(code As System.Text.StringBuilder, typeBlock As Syntax.TypeBlockSyntax, compilation As Compilation, cancellationToken As Threading.CancellationToken, targetTypeInfo As TypeInfo)
        ' Build the complete type hierarchy by traversing up the syntax tree
        Dim typeHierarchy As New List(Of Syntax.TypeBlockSyntax)()
        Dim currentBlock As Syntax.TypeBlockSyntax = typeBlock
        
        ' Collect all parent type blocks in reverse order (from top to bottom)
        While currentBlock IsNot Nothing
            typeHierarchy.Insert(0, currentBlock)
            ' Find the parent type block (not including current)
            Dim parentBlock = currentBlock.Parent
            While parentBlock IsNot Nothing AndAlso Not (TypeOf parentBlock Is Syntax.TypeBlockSyntax)
                parentBlock = parentBlock.Parent
            End While
            
            If parentBlock Is Nothing Then
                ' We've reached the top level
                Exit While
            End If
            
            currentBlock = TryCast(parentBlock, Syntax.TypeBlockSyntax)
        End While
        
        ' Generate the complete type hierarchy with proper nesting
        GenerateNestedTypeHierarchy(code, typeHierarchy, typeBlock, targetTypeInfo, compilation, cancellationToken, 0)
    End Sub

    Private Sub GenerateNestedTypeHierarchy(code As System.Text.StringBuilder, typeHierarchy As List(Of Syntax.TypeBlockSyntax), targetBlock As Syntax.TypeBlockSyntax, targetTypeInfo As TypeInfo, compilation As Compilation, cancellationToken As Threading.CancellationToken, currentIndex As Integer)
        If currentIndex >= typeHierarchy.Count Then Return
        
        Dim currentBlock As Syntax.TypeBlockSyntax = typeHierarchy(currentIndex)
        Dim isTargetType As Boolean = (currentBlock Is targetBlock)
        Dim isLastType As Boolean = (currentIndex = typeHierarchy.Count - 1)
        
        ' Get type information for current block
        Dim currentTypeInfo As TypeInfo
        If isTargetType Then
            currentTypeInfo = targetTypeInfo
        Else
            currentTypeInfo = AnalyzeType(currentBlock, compilation, cancellationToken)
        End If
        
        If currentTypeInfo Is Nothing Then Exit Sub
        ' Generate the type declaration
        Dim typeKeyword As String
        Select Case currentBlock.BlockStatement.Kind()
            Case SyntaxKind.ClassStatement
                typeKeyword = "Class"
            Case SyntaxKind.StructureStatement
                typeKeyword = "Structure"
            Case SyntaxKind.ModuleStatement
                typeKeyword = "Module"
            Case Else
                typeKeyword = "Class"
        End Select

        Dim indent As New String(" "c, currentIndex * 4)
        Dim typeName As String = currentBlock.BlockStatement.Identifier.ValueText

        ' Get access modifier from the original type
        Dim model = compilation.GetSemanticModel(currentBlock.SyntaxTree)
        Dim typeSymbol = model.GetDeclaredSymbol(currentBlock.BlockStatement, cancellationToken)
        Dim accessModifier As String = GetAccessModifier(typeSymbol)

        ' New in 1.0.5: Add a bracket in case it's a keyword like "MyClass" or "Structure"
        code.AppendLine($"{indent}Partial {accessModifier} {typeKeyword} [{typeName}]")
        If isTargetType AndAlso Not isLastType Then
            ' This shouldn't happen - target type should be the last one
            code.AppendLine($"{indent}End {typeKeyword}")
        ElseIf isTargetType AndAlso isLastType Then
            ' Generate the target type with dispose implementation
            'code.AppendLine()
            GenerateDisposeImplementation(code, targetTypeInfo, currentIndex + 1)
            code.AppendLine()
            code.AppendLine($"{indent}End {typeKeyword}")
        ElseIf Not isTargetType AndAlso isLastType Then
            ' This shouldn't happen - non-target type shouldn't be last
            code.AppendLine($"{indent}End {typeKeyword}")
        Else
            ' This is a parent type, so we need to generate nested types recursively
            GenerateNestedTypeHierarchy(code, typeHierarchy, targetBlock, targetTypeInfo, compilation, cancellationToken, currentIndex + 1)
            code.AppendLine($"{indent}End {typeKeyword}")
        End If
    End Sub

    Private Sub GenerateDisposeImplementation(code As System.Text.StringBuilder, typeInfo As TypeInfo, indentLevel As Integer)
        Dim indent As New String(" "c, indentLevel * 4)
        Dim innerIndent As New String(" "c, (indentLevel + 1) * 4)

        Dim isModule As Boolean = typeInfo.TypeConstruct = TypeConstruct.Module
        Dim isStructure As Boolean = typeInfo.TypeConstruct = TypeConstruct.Structure

        If Not isModule Then
            code.AppendLine($"{indent}Implements IDisposable")
            code.AppendLine()
        End If

        If isModule Then
            code.AppendLine($"{indent}''' <summary>")
            code.AppendLine($"{indent}''' Disposes of all fields marked with <c>&lt;DisposeField&gt;</c> in the {typeInfo.TypeName} module.")
            code.AppendLine($"{indent}''' </summary>")
            code.AppendLine($"{indent}Public Sub DisposeModuleFields()")
            code.AppendLine($"{innerIndent}' Dispose both managed and unmanaged resources in a module")
            For Each field In typeInfo.Fields
                If field.IsDisposable Then
                    Dim nullSafeOp = If(field.IsNullable, "?", "")
                    code.AppendLine($"{innerIndent}{typeInfo.TypeName}.{field.FieldName}{nullSafeOp}.Dispose()")
                    code.AppendLine($"{innerIndent}{typeInfo.TypeName}.{field.FieldName} = Nothing")
                End If
            Next field
            code.AppendLine($"{innerIndent}DisposeUnmanagedResources()")
            code.AppendLine($"{indent}End Sub")
            code.AppendLine()
        Else
            Dim assignment = If(isStructure, "", " = False")
            code.AppendLine($"{indent}Private disposedValue As Boolean{assignment}")
            code.AppendLine()

            ' Determine the correct Dispose method signature
            Dim modifier As String
            Dim hasMyBaseDispose As Boolean = False

            If typeInfo.TypeConstruct = TypeConstruct.Structure Then
                modifier = "Private"
            ElseIf typeInfo.HasDisposableBaseClass Then
                modifier = "Protected Overrides"
                hasMyBaseDispose = True
            ElseIf typeInfo.IsSealedClass Then
                modifier = "Private"
            Else
                modifier = "Protected Overridable"
            End If

            code.AppendLine($"{indent}{modifier} Sub Dispose(disposing As Boolean)")
            code.AppendLine($"{innerIndent}If Not disposedValue Then")
            code.AppendLine($"{innerIndent}    If disposing Then")
            code.AppendLine($"{innerIndent}        ' Dispose managed state")
            For Each field In typeInfo.Fields
                If field.IsDisposable Then
                    Dim nullSafeOp = If(field.IsNullable, "?", "")
                    code.AppendLine($"{innerIndent}        Me.{field.FieldName}{nullSafeOp}.Dispose()")
                    code.AppendLine($"{innerIndent}        Me.{field.FieldName} = Nothing")
                End If
            Next field
            code.AppendLine($"{innerIndent}    End If")
            code.AppendLine()
            code.AppendLine($"{innerIndent}    ' Dispose unmanaged resources (like setting large fields to Nothing)")
            code.AppendLine($"{innerIndent}    DisposeUnmanagedResources()")
            code.AppendLine($"{innerIndent}    disposedValue = True")
            code.AppendLine($"{innerIndent}End If")
            If hasMyBaseDispose Then code.AppendLine($"{innerIndent}MyBase.Dispose(disposing)")
            code.AppendLine($"{indent}End Sub")
            code.AppendLine()
        End If

        ' Partial method for unmanaged resources
        code.AppendLine($"{indent}''' <summary>")
        code.AppendLine($"{indent}''' Implement this partial method to dispose of unmanaged resources, such as")
        code.AppendLine($"{indent}''' setting large fields to Nothing.")
        code.AppendLine($"{indent}''' </summary>")
        code.AppendLine($"{indent}Partial Private Sub DisposeUnmanagedResources()")
        code.AppendLine($"{indent}End Sub")
        code.AppendLine()

        If Not isModule Then
            ' Public Dispose implementation
            code.AppendLine($"{indent}Public Sub Dispose() Implements IDisposable.Dispose")
            code.AppendLine($"{innerIndent}' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method")
            code.AppendLine($"{innerIndent}Dispose(disposing:=True)")
            code.AppendLine($"{innerIndent}GC.SuppressFinalize(Me)")
            code.AppendLine($"{indent}End Sub")
            code.AppendLine()

            ' Finalizer (only for classes, not structures)
            If Not isStructure Then
                code.AppendLine($"{indent}Protected Overrides Sub Finalize()")
                code.AppendLine($"{innerIndent}Try")
                code.AppendLine($"{innerIndent}    Dispose(disposing:=False)")
                code.AppendLine($"{innerIndent}Finally")
                code.AppendLine($"{innerIndent}    MyBase.Finalize()")
                code.AppendLine($"{innerIndent}End Try")
                code.AppendLine($"{indent}End Sub")
            End If
        End If
    End Sub

    Private Sub GenerateTypeHierarchy(code As System.Text.StringBuilder, typeInfo As TypeInfo, indentLevel As Integer)
        ' Type declaration
        Dim typeKeyword As String
        Select Case typeInfo.TypeConstruct
            Case TypeConstruct.Class
                typeKeyword = "Class"
            Case TypeConstruct.Structure
                typeKeyword = "Structure"
            Case TypeConstruct.Module
                typeKeyword = "Module"
            Case Else
                typeKeyword = "Class"
        End Select

        Dim isModule As Boolean = typeInfo.TypeConstruct = TypeConstruct.Module
        Dim isStructure As Boolean = typeInfo.TypeConstruct = TypeConstruct.Structure
        Dim indent As New String(" "c, indentLevel * 4)

        ' New in 1.0.5: Add a bracket in case it is a keyword like "MyClass" or "Structure"
        code.AppendLine($"{indent}Partial {typeInfo.AccessModifier} {typeKeyword} [{typeInfo.TypeName}]")
        If Not isModule Then code.AppendLine($"{indent}    Implements IDisposable")
        code.AppendLine()

        If isModule Then
            code.AppendLine($"{indent}    ''' <summary>")
            code.AppendLine($"{indent}    ''' Disposes of all fields marked with <c>&lt;DisposeField&gt;</c> in the {typeInfo.TypeName} module.")
            code.AppendLine($"{indent}    ''' </summary>")
            code.AppendLine($"{indent}    Public Sub DisposeModuleFields()")
            code.AppendLine($"{indent}        ' Dispose both managed and unmanaged resources in a module")
            For Each field In typeInfo.Fields
                If field.IsDisposable Then
                    Dim nullSafeOp = If(field.IsNullable, "?", "")
                    code.AppendLine($"{indent}        {typeInfo.TypeName}.{field.FieldName}{nullSafeOp}.Dispose()")
                    code.AppendLine($"{indent}        {typeInfo.TypeName}.{field.FieldName} = Nothing")
                End If
            Next field
            code.AppendLine($"{indent}        DisposeUnmanagedResources()")
            code.AppendLine($"{indent}    End Sub")
            code.AppendLine()
        Else
            Dim assignment = If(isStructure, "", " = False")
            code.AppendLine($"{indent}    Private disposedValue As Boolean{assignment}")
            code.AppendLine()
            
            ' Determine the correct Dispose method signature based on class inheritance and sealed status
            Dim modifier As String
            Dim hasMyBaseDispose As Boolean = False
            
            If typeInfo.TypeConstruct = TypeConstruct.Structure Then
                modifier = "Private"
            ElseIf typeInfo.HasDisposableBaseClass Then
                modifier = "Protected Overrides"
                hasMyBaseDispose = True
            ElseIf typeInfo.IsSealedClass Then
                modifier = "Private"
            Else
                modifier = "Protected Overridable"
            End If
            
            code.AppendLine($"{indent}    {modifier} Sub Dispose(disposing As Boolean)")
            code.AppendLine($"{indent}        If Not disposedValue Then")
            code.AppendLine($"{indent}            If disposing Then")
            code.AppendLine($"{indent}                ' Dispose managed state")
            For Each field In typeInfo.Fields
                If field.IsDisposable Then
                    Dim nullSafeOp = If(field.IsNullable, "?", "")
                    code.AppendLine($"{indent}                Me.{field.FieldName}{nullSafeOp}.Dispose()")
                    code.AppendLine($"{indent}                Me.{field.FieldName} = Nothing")
                End If
            Next field
            code.AppendLine($"{indent}            End If")
            code.AppendLine()
            code.AppendLine($"{indent}            ' Dispose unmanaged resources (like setting large fields to Nothing)")
            code.AppendLine($"{indent}            DisposeUnmanagedResources()")
            code.AppendLine($"{indent}            disposedValue = True")
            code.AppendLine($"{indent}        End If")
            If hasMyBaseDispose Then code.AppendLine($"{indent}        MyBase.Dispose(disposing)")
            code.AppendLine($"{indent}    End Sub")
            code.AppendLine()
        End If

        ' Partial method for unmanaged resources (will be implemented, not overridden)
        code.AppendLine($"{indent}    ''' <summary>")
        code.AppendLine($"{indent}    ''' Implement this partial method to dispose of unmanaged resources, such as")
        code.AppendLine($"{indent}    ''' setting large fields to Nothing.")
        code.AppendLine($"{indent}    ''' </summary>")
        code.AppendLine($"{indent}    Partial Private Sub DisposeUnmanagedResources()")
        code.AppendLine($"{indent}    End Sub")
        code.AppendLine()

        If Not isModule Then
            ' Public Dispose implementation
            code.AppendLine($"{indent}    Public Sub Dispose() Implements IDisposable.Dispose")
            code.AppendLine($"{indent}        ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method")
            code.AppendLine($"{indent}        Dispose(disposing:=True)")
            code.AppendLine($"{indent}        GC.SuppressFinalize(Me)")
            code.AppendLine($"{indent}    End Sub")
            code.AppendLine()

            ' Finalizer (only for classes, not structures)
            If Not isStructure Then
                code.AppendLine($"{indent}    Protected Overrides Sub Finalize()")
                code.AppendLine($"{indent}        Try")
                code.AppendLine($"{indent}            Dispose(disposing:=False)")
                code.AppendLine($"{indent}        Finally")
                code.AppendLine($"{indent}            MyBase.Finalize()")
                code.AppendLine($"{indent}        End Try")
                code.AppendLine($"{indent}    End Sub")
            End If
        End If

        code.AppendLine($"{indent}End {typeKeyword}")
    End Sub
End Class