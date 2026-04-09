# FieldDisposerVB.Generator

A VB.NET source generator that automatically generates field disposal methods for classes that implement IDisposable. This tool simplifies resource management by automatically creating proper dispose patterns for fields marked with the `<DisposeField>` attribute.

Support for `Structure` with fields marked with `<DisposeField>` is broken through 1.0.0 and 1.0.1. *__To utilize the `<DisposeField>` attribute in VB.NET structures, please use version 1.0.2 or later.__*

> **Notes of Usage:**
>
> If a class/structure contains at least one field marked with `<DisposeField>`, do NOT implement the `IDisposable` interface manually. The source generator will automatically implement this interface and handle the disposal of marked fields.
>
> However, when it is a **module** that contains at least one field marked with `<DisposeField>`, you **must manually call** the `DisposeModuleFields()` method to dispose all the fields marked with this attribute.

## Version 1.0.2 Update

This release introduces significant improvements to field disposal handling and type accessibility:

### New Features
- **Smart Nullability Handling**: The generator now intelligently determines if fields are nullable and generates appropriate disposal code:
  - For nullable fields: Uses safe `?.` operator (`field?.Dispose()`)
  - For non-nullable fields: Will be disposed directly (`field.Dispose()`)
- **Access Modifier Support**: Generated partial types now correctly match the original type's access modifier (Public, Private, Friend, Protected Friend, etc.)
- **Enhanced Type Analysis**: Improved detection of field nullability and type accessibility

### Known Limitations
- **Nested Classes/Structures**: The source generator currently does not support nested classes or structures with fields marked with `<DisposeField>` attribute in them. Only top-level types are processed.

## Version 1.0.1 Update

This release fixes a minor documentation typo in the generated code comments. In version 1.0.0, the XML documentation **incorrectly** stated "Override this partial method" when referring to the `DisposeUnmanagedResources()` partial method. This has been corrected to **"Implement this partial method"** to accurately reflect that partial methods are implemented, not overridden.

> **Important:** The functionality of the package **remains identical** to the previous version (1.0.0). Only the generated XML documentation comment has been corrected for accuracy and professionalism in this version (1.0.1).

## Requirements

- Visual Basic .NET (VB.NET) language version **14.0 or higher**
- .NET Standard 2.0 or higher

## Features

- Automatically generates `IDisposable` implementation
- Creates proper dispose pattern with managed/unmanaged resource handling
- Generates finalizer to ensure cleanup
- Handles nullable field disposal safely
- Provides partial method for custom unmanaged resource disposal
- Supports classes, structures, and modules

## Installation

Install the package via NuGet:

```powershell
Install-Package FieldDisposerVB.Generator
```

Or via .NET CLI:

```bash
dotnet add package FieldDisposerVB.Generator
```

## Usage

### Basic Usage

Mark fields with the `<DisposeField>` attribute to have them automatically disposed:

```vb
<DisposeField>
Private _stream As FileStream

Public Sub New(filePath As String)
    _stream = New FileStream(filePath, FileMode.Open)
End Sub
```

The source generator will automatically create:

- An `IDisposable` interface implementation
- A protected `Dispose(Boolean)` method that disposes marked fields
- A public `Dispose()` method
- A finalizer
- A partial method `DisposeUnmanagedResources()` for custom cleanup

### Complete Example for Classes/Structures

```vb
Imports System.IO

' Mark class as partial to allow source generator to add members
Partial Public Class FileManager
    ' Fields marked with <DisposeField> will be automatically disposed
    <DisposeField>
    Private _inputStream As FileStream
    
    <DisposeField>
    Private _outputStream As MemoryStream
    
    ' Regular field that won't be auto-disposed
    Private _buffer As Byte()
    
    Public Sub New(inputPath As String)
        _inputStream = New FileStream(inputPath, FileMode.Open)
        _outputStream = New MemoryStream()
        _buffer = New Byte(1023) {}
    End Sub
    
    ' The source generator will add the entire dispose pattern here
    ' You can add custom unmanaged resource disposal logic:
    Private Sub DisposeUnmanagedResources()
        ' Add custom disposal logic for unmanaged resources here
        ' For example: setting large arrays to Nothing
        _buffer = Nothing
    End Sub
End Class
```

After compilation, the source generator creates a partial class file that implements the complete dispose pattern:

```vb
' Generated file (FileManager_Disposers.g.vb)
Partial Public Class FileManager
    Implements IDisposable

    Private disposedValue As Boolean = False

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' Dispose managed state
                Me._inputStream?.Dispose()
                Me._inputStream = Nothing
                Me._outputStream?.Dispose()
                Me._outputStream = Nothing
            End If

            ' Dispose unmanaged resources (like setting large fields to Nothing)
            DisposeUnmanagedResources()
            disposedValue = True
        End If
    End Sub

    ''' <summary>
    ''' Implement this partial method to dispose of unmanaged resources, such as
    ''' setting large fields to Nothing.
    ''' </summary>
    Partial Private Sub DisposeUnmanagedResources()
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
        Dispose(disposing:=True)
        GC.SuppressFinalize(Me)
    End Sub

    Protected Overrides Sub Finalize()
        Try
            Dispose(disposing:=False)
        Finally
            MyBase.Finalize()
        End Try
    End Sub
End Class
```

### Further Example for Modules

The source generator also supports modules with the `DisposeModuleFields()` method:

```vb
Imports System.IO

' Modules are supported too
Partial Public Module ResourceManager
    ' Fields marked with <DisposeField> will be automatically disposed via DisposeModuleFields()
    <DisposeField>
    Public CurrentLogStream As FileStream
    
    <DisposeField>
    Private _cache As MemoryStream
    
    ' Regular field - won't be auto-disposed
    Private _tempData As Byte()
    
    Sub Initialize()
        CurrentLogStream = New FileStream("app.log", FileMode.OpenOrCreate)
        _cache = New MemoryStream()
        _tempData = New Byte(2047) {}
    End Sub
    
    ' You can add custom unmanaged resource disposal logic:
    Private Sub DisposeUnmanagedResources()
        ' Add custom disposal logic for unmanaged resources here
        _tempData = Nothing
    End Sub
End Module
```

For modules, the source generator creates:

```vb
' Generated file (ResourceManager_Disposers.g.vb)
Partial Public Module ResourceManager
    ''' <summary>
    ''' Disposes of all fields marked with <c>&lt;DisposeField&gt;</c> in the ResourceManager module.
    ''' </summary>
    Public Sub DisposeModuleFields()
        ' Dispose both managed and unmanaged resources in a module
        ResourceManager.CurrentLogStream?.Dispose()
        ResourceManager.CurrentLogStream = Nothing
        ResourceManager._cache?.Dispose()
        ResourceManager._cache = Nothing
        DisposeUnmanagedResources()
    End Sub

    ''' <summary>
    ''' Implement this partial method to dispose of unmanaged resources, such as
    ''' setting large fields to Nothing.
    ''' </summary>
    Private Sub DisposeUnmanagedResources()
    End Sub
End Module
```

To dispose of module resources, call the generated `DisposeModuleFields()` method:

```vb
' When cleaning up resources in your application
ResourceManager.DisposeModuleFields()
```

## How It Works

The source generator scans your VB.NET code for:

1. Partial classes, structures, or modules
2. Fields marked with the `<DisposeField>` attribute
3. Fields whose types implement `IDisposable`

At compile time, it generates a partial class that implements the standard dispose pattern, ensuring that all marked fields are properly disposed when the containing object is disposed.

## License

This project is licensed under the BSD 3-Clause License. See the [LICENSE](LICENSE) file for details.