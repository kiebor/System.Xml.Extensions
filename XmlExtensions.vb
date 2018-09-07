Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Runtime.Serialization
Imports System.Text
Imports System.Web.Script.Serialization
Imports System.Xml.Serialization


''' <summary>
''' 
''' </summary>
Public Module XmlExtensions
    ''' <summary>
    ''' Returns a new XElement or modifies existing, populated with the specified value.
    ''' </summary>
    ''' <param name="element">The element.</param>
    ''' <param name="value">The value.</param>
    ''' <returns></returns>
    <Extension>
    Function WithValue(element As XElement, value As Object) As XElement
        element.Value = value : Return element
    End Function

    ''' <summary>
    ''' Returns a new XElement or modifies existing, adding a single specified attribute.
    ''' </summary>
    ''' <param name="element">The element.</param>
    ''' <param name="name">The attribute name.</param>
    ''' <param name="value">The value.</param>
    ''' <returns></returns>
    <Extension>
    Function WithAttribute(element As XElement, name As String, value As Object) As XElement
        element.Add(New XAttribute(name, value)) : Return element
    End Function

    ''' <summary>
    ''' Returns a new XElement or modifies existing, adding 1 or more specified attributes.
    ''' </summary>
    ''' <param name="element">The element.</param>
    ''' <param name="attributes">The attribute name.</param>
    ''' <returns></returns>
    <Extension>
    Function WithAttributes(element As XElement, attributes As IEnumerable(Of XAttribute)) As XElement
        element.Add(attributes) : Return element
    End Function

    ''' <summary>
    ''' Returns a new XElement or modifies existing, adding a single specified attribute.
    ''' </summary>
    ''' <param name="element">The element.</param>
    ''' <param name="kvp">The KeyValuePair.</param>
    ''' <returns></returns>
    <Extension>
    Function WithAttribute(element As XElement, kvp As KeyValuePair(Of String, Object)) As XElement
        element.Add(kvp.ToXAttribute) : Return element
    End Function

    ''' <summary>
    ''' Returns a new XElement or modifies existing, adding 1 or more specified attributes.
    ''' </summary>
    ''' <param name="element">The element.</param>
    ''' <param name="dictionary">The Dictionary.</param>
    ''' <returns></returns>
    <Extension>
    Function WithAttributes(element As XElement, dictionary As Dictionary(Of String, Object)) As XElement
        element.Add(dictionary.AsXAttributes) : Return element
    End Function

    ''' <summary>
    ''' Returns a new XElement or modifies existing, populated with the injected XElements.
    ''' </summary>
    ''' <param name="element">The element.</param>
    ''' <param name="insertables">The insertables.</param>
    ''' <returns></returns>
    <Extension>
    Function [From](element As XElement, insertables As IEnumerable(Of XElement)) As XElement
        element.AddRange(insertables) : Return element
    End Function

    ''' <summary>
    ''' Adds the specified range.
    ''' </summary>
    ''' <param name="element">The element.</param>
    ''' <param name="insertables">The insertables.</param>
    <Extension>
    Sub AddRange(element As XElement, insertables As IEnumerable(Of XElement))
        insertables.ToList.ForEach(Sub(e) element.Add(e))
    End Sub

    ''' <summary>
    '''Returns a new XElement or modifies existing, injected with the specified child element.
    ''' </summary>
    ''' <param name="element">The element.</param>
    ''' <param name="insertable">The insertable.</param>
    ''' <returns></returns>
    <Extension>
    Function [With](element As XElement, insertable As XElement) As XElement
        element.Add(insertable) : Return element
    End Function

    ''' <summary>
    ''' Converts a valid xml string structure to XDocument
    ''' </summary>
    ''' <param name="xmlstring">The xmlstring.</param>
    ''' <returns></returns>
    ''' <exception cref="XmlException"></exception>
    <Extension>
    Function AsXDocument(xmlstring As String) As XDocument
        Try
            Return XDocument.Parse(xmlstring)
        Catch ex As Exception
            Throw New XmlException($"Could not Parse string to xml; not a valid xml structure: '{xmlstring}'")
        End Try
    End Function

    ''' <summary>
    ''' Converts Xelement structure to XDocument
    ''' </summary>
    ''' <param name="xml">The XML.</param>
    ''' <returns></returns>
    <Extension>
    Function ToXDocument(xml As XElement) As XDocument
        Return New XDocument(xml)
    End Function

    ''' <summary>
    ''' Converts a string to XElement with value
    ''' </summary>
    ''' <param name="name">The name.</param>
    ''' <param name="value">The value.</param>
    ''' <returns></returns>
    <Extension>
    Function ToXElement(name As String, value As Object) As XElement
        Return New XElement(name, value)
    End Function


    ''' <summary>
    ''' Converts a KeyValuePair to XElement with value
    ''' </summary>
    ''' <param name="kvp">The KeyValuePair.</param>
    ''' <returns></returns>
    <Extension>
    Function ToXElement(kvp As KeyValuePair(Of String, Object)) As XElement
        Return New XElement(kvp.Key, kvp.Value)
    End Function

    ''' <summary>
    ''' Converts a KeyValuePair to XAttribute with value
    ''' </summary>
    ''' <param name="kvp">The KVP.</param>
    ''' <returns></returns>
    <Extension>
    Function ToXAttribute(kvp As KeyValuePair(Of String, Object)) As XAttribute
        Return New XAttribute(kvp.Key, kvp.Value)
    End Function

    ''' <summary>
    ''' Converts a string to XAttribute with value
    ''' </summary>
    ''' <param name="name">The name.</param>
    ''' <param name="value">The value.</param>
    ''' <returns></returns>
    <Extension>
    Function ToXAttribute(name As String, value As Object) As XAttribute
        Return New XAttribute(name, value)
    End Function

    ''' <summary>
    ''' Converts a Dictionary to XElement with value
    ''' </summary>
    ''' <param name="dictionary">The Dictionary.</param>
    ''' <returns></returns>
    <Extension>
    Function AsXElements(dictionary As Dictionary(Of String, Object)) As IEnumerable(Of XElement)
        Dim elements As New List(Of XElement)
        dictionary.ToList().ForEach(Sub(kvp) elements.Add(kvp.ToXElement()))
        Return elements
    End Function


    ''' <summary>
    ''' Converts a Dictionary to XAttributes with value
    ''' </summary>
    ''' <param name="dictionary">The Dictionary.</param>
    ''' <returns></returns>
    <Extension>
    Function AsXAttributes(dictionary As Dictionary(Of String, Object)) As IEnumerable(Of XAttribute)
        Dim attributes As New List(Of XAttribute)
        dictionary.ToList.ForEach(Sub(kvp) attributes.Add(kvp.ToXAttribute))
        Return attributes
    End Function

    ''' <summary>
    ''' Returns a new XDocument or modifies existing, readied with the specified XDeclaration.
    ''' </summary>
    ''' <param name="document">The document.</param>
    ''' <param name="declaration">The declaration.</param>
    ''' <returns></returns>
    <Extension>
    Function WithDeclaration(document As XDocument, declaration As XDeclaration) As XDocument
        document.Declaration = declaration : Return document
    End Function

    ''' <summary>
    ''' Returns a new XDocument or modifies existing, populated from the specified XElement as root.
    ''' </summary>
    ''' <param name="document">The document.</param>
    ''' <param name="element">The element.</param>
    ''' <returns></returns>
    <Extension>
    Function [From](document As XDocument, element As XElement) As XDocument
        document.Insert(element) : Return document
    End Function

    ''' <summary>
    ''' Inserts specified XElement into the defined XDocument.
    ''' </summary>
    ''' <param name="document">The document.</param>
    ''' <param name="element">The element.</param>
    <Extension>
    Sub Insert(document As XDocument, element As XElement)
        If document.Elements.Count < 1 Then
            document.Add(element)
        Else
            document.Root.Add(element)
        End If
    End Sub

    ''' <summary>
    ''' Inserts specified XElement array into the defined XDocument.
    ''' </summary>
    ''' <param name="document">The document.</param>
    ''' <param name="element">The element.</param>
    ''' <exception cref="System.Xml.XmlParseException">
    ''' Can not insert multiple elements @root. This would produce an invalid xml structure.
    ''' or
    ''' Error @root; potentially invalid structure. See inner exception for detail.
    ''' </exception>
    <Extension>
    Sub Insert(document As XDocument, element As IEnumerable(Of XElement))
        If document.Elements.Count < 1 Then
            Throw New XmlParseException("Can not insert multiple elements @root. This would produce an invalid xml structure.")
        Else
            Try
                document.Root.Add(element)
            Catch ex As Exception
                Throw New XmlParseException("Error @root; potentially invalid structure. See inner exception for detail.", ex)
            End Try

        End If

    End Sub

    ''' <summary>
    ''' Returns a new XDocument or modifies existing, injected with the specified root element.
    ''' </summary>
    ''' <param name="document">The document.</param>
    ''' <param name="root">The root.</param>
    ''' <returns></returns>
    <Extension>
    Public Function WithRoot(document As XDocument, Optional root As String = "Root") As XDocument
        If document.Elements.Count < 1 Then document.Insert(New XElement(root)) Else Throw New XmlParseException($"Document already has root element: <{document.Root.Name.ToString}>")
        Return document
    End Function

    ''' <summary>
    ''' Returns a new XDocument or modifies existing, injected with the specified root element.
    ''' </summary>
    ''' <param name="document">The document.</param>
    ''' <param name="root">The root.</param>
    ''' <returns></returns>
    <Extension>
    Public Function WithRoot(document As XDocument, root As XElement) As XDocument
        If document.Elements.Count < 1 Then document.Insert(root) Else Throw New XmlParseException($"Document already has root element: <{document.Root.Name.ToString}>")
        Return document
    End Function

    ''' <summary>
    ''' Returns modified XDocument by adding namespacing at the root element.
    ''' </summary>
    ''' <param name="document">The document.</param>
    ''' <param name="xmlnsUrl">The XMLNS URL.</param>
    ''' <param name="xsdUrl">The XSD URL.</param>
    ''' <param name="xsiUrl">The xsi URL.</param>
    ''' <returns></returns>
    <Extension>
    Function WithNamespace(document As XDocument, xmlnsUrl As String, xsdUrl As String, xsiUrl As String) As XDocument
        Try
            document.Root.Add(
                New XAttribute("xmlns", XNamespace.Get(xmlnsUrl).NamespaceName),
                New XAttribute(XNamespace.Xmlns + "xsd", XNamespace.Get(xsdUrl).NamespaceName),
                New XAttribute(XNamespace.Xmlns + "xsi", XNamespace.Get(xsiUrl).NamespaceName)
            )
            Return document
        Catch ex As Exception
            Throw New XmlParseException("Unable to insert Namespace; potentially missing root element.", ex)
        End Try

    End Function

    ''' <summary>
    ''' Returns modified XDocument by adding namespacing at the root element.
    ''' </summary>
    ''' <param name="document">The document.</param>
    ''' <param name="xmlns">The XMLNS.</param>
    ''' <param name="xsd">The XSD.</param>
    ''' <param name="xsi">The xsi.</param>
    ''' <returns></returns>
    <Extension>
    Function WithNamespace(document As XDocument, xmlns As XNamespace, xsd As XNamespace, xsi As XNamespace) As XDocument
        Try
            document.Root.Add(
                New XAttribute("xmlns", xmlns.NamespaceName),
                New XAttribute(XNamespace.Xmlns + "xsd", xsd.NamespaceName),
                New XAttribute(XNamespace.Xmlns + "xsi", xsi.NamespaceName)
            )
            Return document
        Catch ex As Exception
            Throw New XmlParseException("Unable to insert Namespace; potentially missing root element.", ex)
        End Try
    End Function

    ''' <summary>
    ''' Serialize T (from Type) to XString object
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="value">The value.</param>
    ''' <returns></returns>
    ''' <exception cref="System.Xml.XmlParseException">An error occurred; see inner exception for details.</exception>
    <Extension>
    Public Function ToXString(Of T)(value As T) As XString
        If value Is Nothing Then Return New XString(String.Empty)

        Try
            Dim xmlserializer = New XmlSerializer(GetType(T)), stringWriter = New StringWriter()
            Using writer = XmlWriter.Create(stringWriter)
                xmlserializer.Serialize(writer, value)
                Return New XString(stringWriter.ToString())

            End Using

        Catch ex As Exception
            Throw New XmlParseException("An error occurred; see inner exception for details.", ex)
        End Try

    End Function

    ''' <summary>
    ''' Converts XDocument to XString
    ''' </summary>
    ''' <param name="doc">The document.</param>
    ''' <returns></returns>
    <Extension>
    Function ToXString(doc As XDocument) As XString
        Return New XString(doc.ToString())
    End Function

    ''' <summary>
    ''' Converts XElement to XString
    ''' </summary>
    ''' <param name="element">The element.</param>
    ''' <returns></returns>
    <Extension>
    Function ToXString(element As XElement) As XString
        Return New XString(element.ToString())
    End Function

    ''' <summary>
    ''' Returns a new XDocument from Json input; invalid or badly formatted json will throw XmlParseException
    ''' </summary>
    ''' <param name="doc">The document.</param>
    ''' <param name="jsonstring">The jsonstring.</param>
    ''' <returns></returns>
    ''' <exception cref="System.Xml.XmlParseException"></exception>
    <Extension>
    Function FromJson(doc As XDocument, jsonstring As String) As XDocument
        Try
            Return XDocument.Load(Json.JsonReaderWriterFactory.CreateJsonReader(Encoding.ASCII.GetBytes(jsonstring), New XmlDictionaryReaderQuotas()))
        Catch ex As Exception
            Throw New XmlParseException($"Could not parse json input. Potentially bad sub structure; see inner exception for details.", ex)
        End Try

    End Function

    ''' <summary>
    ''' Converts an Xml structured string to Json formatted string
    ''' </summary>
    ''' <param name="xmlString">The XML string.</param>
    ''' <returns></returns>
    ''' <exception cref="System.Xml.XmlParseException">Could not parse xml input; see inner exception for details.</exception>
    <Extension>
    Function ToJson(xmlString As String) As String
        Try
            Return Jsonify(XElement.Parse(xmlString).ToRawData())
        Catch ex As Exception
            Throw New XmlParseException("Could not parse xml input; see inner exception for details.", ex)
        End Try

    End Function

    ''' <summary>
    ''' Converts XElement to Json formatted string
    ''' </summary>
    ''' <param name="element">The element.</param>
    ''' <returns></returns>
    <Extension>
    Function ToJson(element As XElement) As String
        Return Jsonify(element.ToRawData())
    End Function


    ''' <summary>
    ''' Converts XDocument to Json formatted string
    ''' </summary>
    ''' <param name="document">The element.</param>
    ''' <returns></returns>
    <Extension>
    Function ToJson(document As XDocument) As String
        Try
            Return Jsonify(document.Root.ToRawData())
        Catch ex As Exception
            Throw New XmlParseException("Potentially invalid structure @root; see inner exception for detail.", ex)
        End Try

    End Function

    Private Function Jsonify([structure] As Dictionary(Of String, Object)) As String
        Return (New JavaScriptSerializer()).Serialize([structure])
    End Function

    ''' <summary>
    ''' Converts XElement to raw data Dictionary
    ''' </summary>
    ''' <param name="xml">The XML.</param>
    ''' <returns></returns>
    <Extension>
    Public Function ToRawData(xml As XElement) As Dictionary(Of String, Object)
        Dim attributes = xml.Attributes().ToDictionary(Function(d) d.Name.LocalName, Function(d) DirectCast(d.Value, Object))
        If xml.HasElements Then
            attributes.Add("_value", xml.Elements().Select(Function(e) e.ToRawData()))
        ElseIf Not xml.IsEmpty Then
            attributes.Add("_value", xml.Value)
        End If

        Return New Dictionary(Of String, Object) From {
                {xml.Name.LocalName, attributes}
            }
    End Function

    ''' <summary>
    ''' Writes the XDocument to StringBuilder; must be converted to usable type after return.
    ''' </summary>
    ''' <param name="document">The document.</param>
    ''' <param name="omitDeclaration">if set to <c>true</c> [omit declaration].</param>
    ''' <param name="preserveIndentation">if set to <c>true</c> [preserve indentation].</param>
    ''' <returns></returns>
    <Extension>
    Public Function WriteTo(document As XDocument, omitDeclaration As Boolean, preserveIndentation As Boolean) As StringBuilder
        Dim sb As New StringBuilder(), xws As New XmlWriterSettings With {
            .OmitXmlDeclaration = omitDeclaration,
            .Indent = preserveIndentation
        }

        Using xw As XmlWriter = XmlWriter.Create(sb, xws)
            document.WriteTo(xw) : Return sb ''.ToString
        End Using

    End Function

    ''' <summary>
    ''' Returns a <see cref="System.String" /> that represents this instance; wraps an XmlWriter return.
    ''' </summary>
    ''' <param name="document">The document.</param>
    ''' <param name="omitDeclaration">if set to <c>true</c> [omit declaration].</param>
    ''' <param name="preserveIndentation">if set to <c>true</c> [preserve indentation].</param>
    ''' <returns>
    ''' A <see cref="System.String" /> that represents this instance.
    ''' </returns>
    <Extension>
    Public Function ToString(document As XDocument, omitDeclaration As Boolean, preserveIndentation As Boolean) As String
        Return document.WriteTo(omitDeclaration, preserveIndentation).ToString()
    End Function

    ''' <summary>
    ''' Returns a <see cref="System.Byte" /> array that represents this instance; wraps an XmlWriter return.
    ''' </summary>
    ''' <param name="document">The document.</param>
    ''' <param name="omitDeclaration">if set to <c>true</c> [omit declaration].</param>
    ''' <param name="preserveIndentation">if set to <c>true</c> [preserve indentation].</param>
    ''' <returns></returns>
    <Extension>
    Public Function ToBytes(document As XDocument, omitDeclaration As Boolean, preserveIndentation As Boolean) As Byte()
        Return Encoding.ASCII.GetBytes(document.WriteTo(omitDeclaration, preserveIndentation).ToString())
    End Function

End Module

''' <summary>
''' XString is a contianer object for securely holding xml converted to string
''' </summary>
Public Class XString
    ''' <summary>
    ''' Gets or sets the content.
    ''' </summary>
    ''' <value>
    ''' The content.
    ''' </value>
    Public Property Content As String

    ''' <summary>
    ''' Initializes a new instance of the <see cref="XString"/> class.
    ''' </summary>
    ''' <param name="xmlstring">The xmlstring.</param>
    Sub New(xmlstring As String)
        Content = xmlstring
    End Sub

End Class

''' <summary>
''' 
''' </summary>
''' <seealso cref="System.Exception" />
Public Class XmlParseException : Inherits Exception
    ''' <summary>
    ''' Initializes a new instance of the <see cref="XmlParseException"/> class.
    ''' </summary>
    ''' <param name="msg">The MSG.</param>
    Sub New(msg As String)
        MyBase.New(msg)

    End Sub

    ''' <summary>
    ''' Initializes a new instance of the <see cref="XmlParseException"/> class.
    ''' </summary>
    ''' <param name="msg">The MSG.</param>
    ''' <param name="ex">The ex.</param>
    Sub New(msg As String, ex As Exception)
        MyBase.New(msg, ex)
    End Sub

End Class
