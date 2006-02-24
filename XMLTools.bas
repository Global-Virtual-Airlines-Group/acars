Attribute VB_Name = "XMLTools"
Option Explicit

Public Function FormatDateTime(dt As Date, fmt As String) As String
    Dim tmpDate As String

    tmpDate = Replace(Format(dt, fmt), config.DateSeparator, "/")
    FormatDateTime = Replace(tmpDate, config.TimeSeparator, ":")
End Function

Public Function ParseDateTime(dt As String) As Date
    Dim tmpDate As String
    
    tmpDate = Replace(dt, "/", config.DateSeparator)
    tmpDate = Replace(tmpDate, ":", config.TimeSeparator)
    ParseDateTime = CDate(tmpDate)
End Function

Public Function FormatNumber(num As Double, fmt As String) As String
    FormatNumber = Replace(Format(num, fmt), config.DecimalSeparator, ".")
End Function

Public Function ParseNumber(num As String) As Double
    ParseNumber = CDbl(Replace(num, ".", config.DecimalSeparator))
End Function

Public Sub AddXMLField(objNode As IXMLDOMElement, Name As String, value As String, Optional doCDATA As Boolean = True)
    Dim objNewNode As IXMLDOMElement
    Dim objCDATA As IXMLDOMCDATASection

    Set objNewNode = objNode.ownerDocument.createNode(NODE_ELEMENT, Name, "")
    If doCDATA Then
        Set objCDATA = objNode.ownerDocument.createCDATASection(value)
        objNewNode.appendChild objCDATA
    Else
        objNewNode.Text = value
    End If
    
    objNode.appendChild objNewNode
End Sub

Public Function getChild(node As IXMLDOMNode, Name As String, Optional defVal As String)
    Dim e As IXMLDOMNode

    Set e = node.selectSingleNode(Name)
    If (e Is Nothing) Then
        getChild = defVal
    Else
        getChild = Trim(e.Text)
    End If
End Function

Public Function getAttr(node As IXMLDOMNode, Name As String, Optional defVal As String)
    Dim e As IXMLDOMAttribute
    
    Set e = node.Attributes.getNamedItem(Name)
    If (e Is Nothing) Then
        getAttr = defVal
    Else
        getAttr = Trim(e.Text)
    End If
End Function
