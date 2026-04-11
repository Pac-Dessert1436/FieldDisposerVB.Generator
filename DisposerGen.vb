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
        For Each typeBlock In typeBlocks
            If typeBlock IsNot Nothing Then
                Dim typeInfo = AnalyzeType(typeBlock, compilation, source.CancellationToken)
                If typeInfo IsNot Nothing AndAlso typeInfo.Fields.Count > 0 Then
                    Dim sourceCode = GenerateDisposeCode(typeInfo)
                    
                    ' Generate unique file name for nested types
                    Dim fileName As String = typeInfo.TypeName
                    If typeInfo.IsNested Then
                        ' For nested types, include the full hierarchy in the file name
                        Dim model = compilation.GetSemanticModel(typeBlock.SyntaxTree)
                        Dim typeSymbol = model.GetDeclaredSymbol(typeBlock.BlockStatement, source.CancellationToken)
                        If typeSymbol IsNot Nothing Then
                            fileName = typeSymbol.ToDisplayString().Replace(".", "_").Replace("+", "_")
                        End If
                    End If
                    
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

        code.AppendLine($"Partial {typeInfo.AccessModifier} {typeKeyword} {typeInfo.TypeName}")
        If Not isModule Then code.AppendLine("    Implements IDisposable")
        code.AppendLine()

        If isModule Then
            code.AppendLine("    ''' <summary>")
            code.AppendLine($"    ''' Disposes of all fields marked with <c>&lt;DisposeField&gt;</c> in the {typeInfo.TypeName} module.")
            code.AppendLine("    ''' </summary>")
            code.AppendLine($"    Public Sub DisposeModuleFields()")
            code.AppendLine("        ' Dispose both managed and unmanaged resources in a module")
            For Each field In typeInfo.Fields
                If field.IsDisposable Then
                    Dim nullSafeOp = If(field.IsNullable, "?", "")
                    code.AppendLine($"        {typeInfo.TypeName}.{field.FieldName}{nullSafeOp}.Dispose()")
                    code.AppendLine($"        {typeInfo.TypeName}.{field.FieldName} = Nothing")
                End If
            Next field
            code.AppendLine("        DisposeUnmanagedResources()")
            code.AppendLine("    End Sub")
            code.AppendLine()
        Else
            Dim assignment = If(isStructure, "", " = False")
            code.AppendLine($"    Private disposedValue As Boolean{assignment}")
            code.AppendLine()
            Dim modifier = If(isStructure, "Private", "Protected Overridable")
            code.AppendLine($"    {modifier} Sub Dispose(disposing As Boolean)")
            code.AppendLine("        If Not disposedValue Then")
            code.AppendLine("            If disposing Then")
            code.AppendLine("                ' Dispose managed state")
            For Each field In typeInfo.Fields
                If field.IsDisposable Then
                    Dim nullSafeOp = If(field.IsNullable, "?", "")
                    code.AppendLine($"                Me.{field.FieldName}{nullSafeOp}.Dispose()")
                    code.AppendLine($"                Me.{field.FieldName} = Nothing")
                End If
            Next field
            code.AppendLine("            End If")
            code.AppendLine()
            code.AppendLine("            ' Dispose unmanaged resources (like setting large fields to Nothing)")
            code.AppendLine("            DisposeUnmanagedResources()")
            code.AppendLine("            disposedValue = True")
            code.AppendLine("        End If")
            code.AppendLine("    End Sub")
            code.AppendLine()
        End If

        ' Partial method for unmanaged resources (will be implemented, not overridden)
        code.AppendLine("    ''' <summary>")
        code.AppendLine("    ''' Implement this partial method to dispose of unmanaged resources, such as")
        code.AppendLine("    ''' setting large fields to Nothing.")
        code.AppendLine("    ''' </summary>")
        code.AppendLine("    Partial Private Sub DisposeUnmanagedResources()")
        code.AppendLine("    End Sub")
        code.AppendLine()

        If Not isModule Then
            ' Public Dispose implementation
            code.AppendLine("    Public Sub Dispose() Implements IDisposable.Dispose")
            code.AppendLine("        ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method")
            code.AppendLine("        Dispose(disposing:=True)")
            code.AppendLine("        GC.SuppressFinalize(Me)")
            code.AppendLine("    End Sub")
            code.AppendLine()

            ' Finalizer (cannot use MyBase in structures)
            code.AppendLine($"    Protected Overrides Sub Finalize()")
            If isStructure Then
                code.AppendLine("        Dispose(disposing:=False)")
            Else
                code.AppendLine("        Try")
                code.AppendLine("            Dispose(disposing:=False)")
                code.AppendLine("        Finally")
                code.AppendLine("            MyBase.Finalize()")
                code.AppendLine("        End Try")  
            End If
            code.AppendLine("    End Sub")
            code.AppendLine()
        End If

        code.AppendLine($"End {typeKeyword}")

        If Not String.IsNullOrEmpty(typeInfo.ContainingNamespace) Then
            code.AppendLine()
            code.AppendLine("End Namespace")
        End If

        Return code.ToString()
    End Function
End Class