Option Explicit On

Imports Microsoft.VisualBasic

Module Module1

    Public Function CTOD(ByVal cFecha As String) As Date

        Dim nDia, nMes, nYear As Integer

        nDia = Val(Right(cFecha, 2))
        nMes = Val(Mid(cFecha, 5, 2))
        nYear = Val(Left(cFecha, 4))

        CTOD = DateSerial(nYear, nMes, nDia)

    End Function

    Public Function DTOC(ByVal dFecha As Date) As String

        Dim cDia, cMes, cYear, sFecha As String

        sFecha = dFecha.ToShortDateString

        cDia = Left(sFecha, 2)
        cMes = Mid(sFecha, 4, 2)
        cYear = Right(sFecha, 4)

        DTOC = cYear & cMes & cDia

    End Function

End Module
