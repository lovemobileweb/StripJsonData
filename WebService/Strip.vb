Imports System.IO
Imports Newtonsoft.Json

Public Class Strip
    Public Function Run(jsonString As String)
        Dim sr = New StringReader(jsonString)
        Dim jsonReader = New JsonTextReader(sr)
        Dim sb = New StringBuilder()
        Dim sw = New StringWriter(sb)
        Dim jsonWriter = New JsonTextWriter(sw)
        While jsonReader.Read()
            If jsonReader.Value IsNot Nothing Then
                If jsonReader.TokenType = JsonToken.PropertyName Then
                    If jsonReader.Value = "authenticationResultCode" Or
                        jsonReader.Value = "brandLogoUri" Or
                        jsonReader.Value = "copyright" Or
                        jsonReader.Value = "__type" Or
                        jsonReader.Value = "traceId" Then
                        ReadValue(jsonReader)
                    ElseIf jsonReader.Value = "bbox" Then
                        ReadArray(jsonReader)
                    Else
                        jsonWriter.WritePropertyName(jsonReader.Value)
                    End If
                ElseIf jsonReader.TokenType = JsonToken.Comment Then
                    jsonWriter.WriteComment(jsonReader.Value)
                Else
                    jsonWriter.WriteValue(jsonReader.Value)
                End If
            Else
                If jsonReader.TokenType = JsonToken.StartObject Then
                    jsonWriter.WriteStartObject()
                ElseIf jsonReader.TokenType = JsonToken.StartArray Then
                    jsonWriter.WriteStartArray()
                ElseIf jsonReader.TokenType = JsonToken.EndObject Then
                    jsonWriter.WriteEndObject()
                ElseIf jsonReader.TokenType = JsonToken.EndArray Then
                    jsonWriter.WriteEndArray()
                End If
            End If
        End While
        jsonReader.Close()
        jsonWriter.Close()
        Run = sb.ToString()
    End Function

    Private Sub ReadArray(reader As JsonTextReader)
        Dim deep As Integer = 0
        While reader.Read()
            If reader.TokenType = JsonToken.StartArray Then
                deep += 1
            ElseIf reader.TokenType = JsonToken.EndArray Then
                deep -= 1
                If deep <= 0 Then
                    Exit While
                End If
            End If
        End While
    End Sub

    Private Sub ReadValue(reader As JsonReader)
        reader.Read()
    End Sub
End Class
