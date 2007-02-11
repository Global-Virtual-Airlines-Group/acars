Attribute VB_Name = "XMLTools"
Option Explicit

Public Function I18nOutputDateTime(dt As UTCDate, Optional fmt As String = "mm/dd/yyyy hh:nn:ss") As String
    Dim tmpDate As String

    tmpDate = Replace(Format(dt.UTCTime, fmt), config.DateSeparator, "/")
    tmpDate = Replace(tmpDate, config.TimeSeparator, ":")
    I18nOutputDateTime = tmpDate & "." & CStr(dt.MilliSeconds)
End Function

Public Function I18nDateTime(dt As String) As String
    Dim tmpDate As String
    
    tmpDate = Replace(dt, "/", config.DateSeparator)
    I18nDateTime = Replace(tmpDate, ":", config.TimeSeparator)
End Function

Public Function FormatNumber(num As Double, fmt As String) As String
    FormatNumber = Replace(Format(num, fmt), config.DecimalSeparator, ".")
End Function

Public Function ParseNumber(num As String) As Double
    ParseNumber = CDbl(Replace(num, ".", config.DecimalSeparator))
End Function

Public Sub AddXMLField(objNode As IXMLDOMElement, name As String, value As String, Optional doCDATA As Boolean = True)
    Dim objNewNode As IXMLDOMElement
    Dim objCDATA As IXMLDOMCDATASection

    Set objNewNode = objNode.ownerDocument.createNode(NODE_ELEMENT, name, "")
    If doCDATA Then
        Set objCDATA = objNode.ownerDocument.createCDATASection(value)
        objNewNode.appendChild objCDATA
    Else
        objNewNode.Text = value
    End If
    
    objNode.appendChild objNewNode
End Sub

Public Sub AddXMLChild(objNode As IXMLDOMElement, name As String, subName As String, value As String, Optional doCDATA As Boolean = False)
    Dim subNode As IXMLDOMElement
    
    Set subNode = objNode.ownerDocument.createNode(NODE_ELEMENT, name, "")
    AddXMLField subNode, subName, value, doCDATA
    objNode.appendChild subNode
End Sub

Public Function getChild(node As IXMLDOMNode, name As String, Optional defVal As String)
    Dim e As IXMLDOMNode

    Set e = node.selectSingleNode(name)
    If (e Is Nothing) Then
        getChild = defVal
    Else
        getChild = Trim(e.Text)
    End If
End Function

Public Function getAttr(node As IXMLDOMNode, name As String, Optional defVal As String)
    Dim e As IXMLDOMAttribute
    
    Set e = node.Attributes.getNamedItem(name)
    If (e Is Nothing) Then
        getAttr = defVal
    Else
        getAttr = Trim(e.Text)
    End If
End Function
