Imports System.Runtime.InteropServices

Public Class MarshalHelper
    Public Shared Function GetActiveObject(
        ByVal progId As String,
        ByVal Optional throwOnError As Boolean = False
        ) As Object

        If progId Is Nothing Then
            Throw New ArgumentNullException(NameOf(progId))
        End If

        Dim clsid = Nothing
        Dim hr = CLSIDFromProgIDEx(progId, clsid)

        If hr < 0 Then
            If throwOnError Then
                System.Runtime.InteropServices.Marshal.ThrowExceptionForHR(hr)
            End If
            Return Nothing
        End If

        Dim obj As Object = Nothing
        hr = GetActiveObject(clsid, IntPtr.Zero, obj)

        If hr < 0 Then
            If throwOnError Then
                System.Runtime.InteropServices.Marshal.ThrowExceptionForHR(hr)
            End If
            Return Nothing
        End If

        Return obj
    End Function

    <DllImport("ole32")>
    Private Shared Function CLSIDFromProgIDEx(
        <MarshalAs(UnmanagedType.LPWStr)> ByVal lpszProgID As String,
        <Out> ByRef lpclsid As Guid
        ) As Integer
    End Function

    <DllImport("oleaut32")>
    Private Shared Function GetActiveObject(
        <MarshalAs(UnmanagedType.LPStruct)> ByVal rclsid As Guid,
        ByVal pvReserved As IntPtr,
        <Out> <MarshalAs(UnmanagedType.IUnknown)> ByRef ppunk As Object
        ) As Integer
    End Function

End Class
