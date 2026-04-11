# FieldDisposerVB.Generator

A VB.NET source generator that automatically generates field disposal methods for classes that implement `IDisposable`. This tool simplifies resource management by automatically creating proper dispose patterns for fields marked with the `<DisposeField>` attribute.

Support for `Structure` with fields marked with `<DisposeField>` is broken through 1.0.0 and 1.0.3 (see [Notes on Broken Versions](#notes-on-broken-versions)). *__To utilize the `<DisposeField>` attribute in VB.NET structures, please use version 1.0.4 or later.__*

> **Notes of Usage:**
>
> If a class/structure contains at least one field marked with `<DisposeField>`, do NOT implement the `IDisposable` interface manually. The source generator will automatically implement this interface and handle the disposal of marked fields.
>
> However, when it is a **module** that contains at least one field marked with `<DisposeField>`, you **must manually call** the `DisposeModuleFields()` method to dispose all the fields marked with this attribute.

## Notes on Broken Versions

Versions 1.0.2 through 1.0.3 are broken due to the following bugs:
- Structures incorrectly used `MyBase.Finalize()` which is not allowed in VB.NET
- In DisposableEvent classes (not inheritable), disposal methods are incorrectly incorrectly declared as `Protected Overridable` instead of `Private`

__DO NOT use 1.0.2 and 1.0.3.__ Please use version 1.0.4 or later for full functionality and bug fixes.

## Working Features in Version 1.0.4

- **Critical Fix**: Fixed compilation errors in structures that incorrectly used `MyBase.Finalize()`
- **Disposable Event Fix**: Corrected disposal method from `Protected Overridable` to `Private` in DisposableEvent classes
- **Nested Class Support**: Added support for nested classes and structures with fields marked with `<DisposeField>` attribute
- **Improved File Naming**: Generated source files for nested types now use unique names based on full type hierarchy
- **Disposable Event Classes**: Added reusable disposable event classes with proper IDisposable implementation (0-5 parameters)
- **Structure Support**: Properly supports structures with fields marked with `<DisposeField>` attribute
- **Smart Nullability Handling**: Generates appropriate disposal code for nullable vs. non-nullable fields
- **Access Modifier Support**: Generated types correctly match original type's access modifier
- **Documentation Fix**: Corrected XML documentation comment from "Override" to "Implement" for partial methods (from version 1.0.1)

## Known Limitations

- **Protected Access**: Classes and structures cannot have protected access modifiers in VB.NET, and the generator properly enforces this language constraint.

## Disposable Event Classes Examples

The generator includes bonus disposable event classes that automatically handle event handler cleanup. These classes implement `IDisposable` and can be used for managing events that need proper disposal.

### Basic Usage Example

```vb
Using simpleEvent As New DisposableEvent
    ' Subscribe & publish the event
    simpleEvent.Subscribe(Sub() Console.WriteLine("Event fired!"))
    simpleEvent.Publish()    
End Using  ' Event handlers automatically cleaned up
```

### Generic Event Example

```vb
Using stringEvent As New DisposableEvent(Of String)
    ' Subscribe & publish the event with parameter
    stringEvent.Subscribe(Sub(msg) Console.WriteLine($"Message: {msg}"))
    stringEvent.Publish("Hello World")
End Using  ' Event handlers automatically cleaned up
```

### Multiple Parameter Example

```vb
Using multiEvent As New DisposableEvent(Of String, Integer)
    ' Subscribe & publish the event with multiple parameters
    multiEvent.Subscribe(Sub(name, age) Console.WriteLine($"Name: {name}, Age: {age}"))
    multiEvent.Publish("John Doe", 30)
End Using  ' Event handlers automatically cleaned up
```

### Real-World Usage Pattern with `<DisposeField>`

```vb
Public Class EventManager
    <DisposeField> Private _userEvent As New DisposableEvent(Of String)
    <DisposeField> Private _dataEvent As New DisposableEvent(Of String, Integer)
    
    Public Sub RegisterUserHandler(handler As Action(Of String))
        _userEvent.Subscribe(handler)
    End Sub
    
    Public Sub RegisterDataHandler(handler As Action(Of String, Integer))
        _dataEvent.Subscribe(handler)
    End Sub
    
    Public Sub NotifyUser(userName As String)
        _userEvent.Publish(userName)
    End Sub
    
    Public Sub NotifyData(dataName As String, dataValue As Integer)
        _dataEvent.Publish(dataName, dataValue)
    End Sub
    
    ' No need for manual cleanup; `Dispose` method automatically implemented
End Class
```

### Benefits of Using DisposableEvent Classes

1. **Automatic Cleanup**: Event handlers are automatically removed when disposed
2. **Memory Safety**: Prevents memory leaks from forgotten event handlers
3. **Thread-Safe**: Uses proper disposal pattern with finalizer
4. **Multiple Signatures**: Supports events with 0 to 5 parameters
5. **Generic Support**: Type-safe event parameters

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