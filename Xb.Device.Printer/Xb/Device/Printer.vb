Option Strict On

Imports System.Drawing.Printing


''' <summary>
''' �v�����^���^���֘A�֐��ێ��N���X
''' </summary>
''' <remarks></remarks>
Partial Public Class Printer

    ''' <summary>
    ''' �C���X�g�[���ς݃v�����^�����X�g���擾����B
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetNames() As List(Of String)

        Dim result As List(Of String) = New List(Of String)()
        For Each prn As String In System.Drawing.Printing.PrinterSettings.InstalledPrinters
            result.Add(prn)
        Next

        Return result

    End Function


    ''' <summary>
    ''' �C���X�g�[���ς݃v�����^�ʂ̖{�N���X�C���X�^���X���擾����B
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetList() As Dictionary(Of String, Device.Printer)

        Dim result As Dictionary(Of String, Device.Printer) = New Dictionary(Of String, Printer)()
        For Each name As String In GetNames()
            result.Add(name, New Device.Printer(name))
        Next

        Return result

    End Function



    Private ReadOnly _printerName As String

    '�Ή��p���T�C�Y
    Private ReadOnly _paperSizes As Dictionary(Of String, System.Drawing.Printing.PaperSize)
    Private ReadOnly _paperNames As List(Of String)
    Private ReadOnly _paperKinds As List(Of System.Drawing.Printing.PaperKind)

    '�v�����^��̃g���C
    Private ReadOnly _trays As Dictionary(Of String, System.Drawing.Printing.PaperSource)
    Private ReadOnly _trayNames As List(Of String)


    ''' <summary>
    ''' �v�����^��
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property PrinterName() As String
        Get
            Return Me._printerName
        End Get
    End Property


    ''' <summary>
    ''' �Ή��p���T�C�Y�I�u�W�F�N�g���X�g
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property PaperSizes() As Dictionary(Of String, System.Drawing.Printing.PaperSize)
        Get
            Return Me._paperSizes
        End Get
    End Property


    ''' <summary>
    ''' �Ή��p���T�C�Y�̃T�C�Y�����X�g
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property PaperNames() As List(Of String)
        Get
            Return Me._paperNames
        End Get
    End Property


    ''' <summary>
    ''' �Ή��p���T�C�Y�̃T�C�Y�탊�X�g
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property PaperKinds() As List(Of System.Drawing.Printing.PaperKind)
        Get
            Return Me._paperKinds
        End Get
    End Property


    ''' <summary>
    ''' �p���g���C�I�u�W�F�N�g���X�g
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Trays() As Dictionary(Of String, System.Drawing.Printing.PaperSource)
        Get
            Return Me._trays
        End Get
    End Property


    ''' <summary>
    ''' �p���g���C�����X�g
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property TrayNames() As List(Of String)
        Get
            Return Me._trayNames
        End Get
    End Property


    ''' <summary>
    ''' �R���X�g���N�^
    ''' </summary>
    ''' <param name="name"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal name As String)

        If (Not Device.Printer.GetNames().Contains(name)) Then
            Xb.Util.Out("Xb.Device.Printer.Constructor: �n���l�v�����^���ɊY������v�����^�����o�ł��܂���B")
            Throw New ArgumentException("Xb.Device.Printer.Constructor: �n���l�v�����^���ɊY������v�����^�����o�ł��܂���B")
        End If

        Me._printerName = name
        Me._paperSizes = New Dictionary(Of String, PaperSize)()
        Me._paperNames = New List(Of String)()
        Me._paperKinds = New List(Of PaperKind)()
        Me._trays = New Dictionary(Of String, PaperSource)()
        Me._trayNames = New List(Of String)()

        Dim settings As System.Drawing.Printing.PageSettings = New System.Drawing.Printing.PageSettings()
        settings.PrinterSettings.PrinterName = name

        For Each size As System.Drawing.Printing.PaperSize In settings.PrinterSettings.PaperSizes
            Me._paperSizes.Add(size.PaperName, size)
            Me._paperNames.Add(size.PaperName)
            Me._paperKinds.Add(size.Kind)
        Next

        For Each source As System.Drawing.Printing.PaperSource In settings.PrinterSettings.PaperSources
            Me._trays.Add(source.SourceName, source)
            Me._trayNames.Add(source.SourceName)
        Next

    End Sub

End Class
