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
    End Class

    Private Class FieldInfo
        Public Property FieldName As String
        Public Property FieldType As String
        Public Property IsDisposable As Boolean
        Public Property IsNullable As Boolean
    End Class

    Public Sub Initialize(context As IGIC) Implements IIncrementalGenerator.Initialize
        ' Register the attribute source
        context.RegisterPostInitializationOutput(Sub(ctx)
                                                     ctx.AddSource("DisposeFieldAttribute.g.vb", GenerateAttributeCode())
                                                 End Sub)

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

                ' Get the containing type block
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

        context.RegisterSourceOutput(compilation, Sub(spc, tup)
                                                      GenerateSource(tup.Left, tup.Right, spc)
                                                  End Sub)
    End Sub

    Private Function GenerateAttributeCode() As String
        Return "Imports System

<AttributeUsage(AttributeTargets.Field, AllowMultiple:=False)>
Public NotInheritable Class DisposeFieldAttribute
    Inherits Attribute

    Public Sub New()
    End Sub
End Class"
    End Function

    Private Sub GenerateSource(compilation As Compilation, typeBlocks As ImmutableArray(Of Syntax.TypeBlockSyntax), source As SourceProductionContext)
        For Each typeBlock In typeBlocks
            If typeBlock IsNot Nothing Then
                Dim typeInfo = AnalyzeType(typeBlock, compilation, source.CancellationToken)
                If typeInfo IsNot Nothing AndAlso typeInfo.Fields.Count > 0 Then
                    Dim sourceCode = GenerateDisposeCode(typeInfo)
                    source.AddSource($"{typeInfo.TypeName}_Disposers.g.vb", sourceCode)
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
        Const DEFAULT_MODIFIER = "Friend"
        If typeSymbol Is Nothing Then Return DEFAULT_MODIFIER

        ' Check the declared accessibility
        Select Case typeSymbol.DeclaredAccessibility
            Case Accessibility.Private
                Return "Private"
            Case Accessibility.Protected
                Return "Protected"
            Case Accessibility.Internal
                Return "Friend"
            Case Accessibility.ProtectedOrInternal
                Return "Protected Friend"
            Case Accessibility.ProtectedAndInternal
                Return "Private Protected"
            Case Accessibility.Public
                Return "Public"
            Case Else
                Return DEFAULT_MODIFIER
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

            ' Finalizer
            code.AppendLine($"    Protected Overrides Sub Finalize()")
            code.AppendLine("        Try")
            code.AppendLine("            Dispose(disposing:=False)")
            code.AppendLine("        Finally")
            code.AppendLine("            MyBase.Finalize()")
            code.AppendLine("        End Try")
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