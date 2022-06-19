'References required:
'UIAutomationClient

Option Explicit

    Private Declare PtrSafe Function MessageBoxW Lib "User32" (ByVal hWnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub test()

    Dim uiAuto As New CUIAutomation
    Dim uiCond_Type As IUIAutomationCondition
    Dim uiCond_Name As IUIAutomationCondition
    Dim uiCond_And As IUIAutomationAndCondition
    Dim uiDesktop As IUIAutomationElement
    Dim uiApp As IUIAutomationElement
    Dim uiAppPattern As IUIAutomationWindowPattern
    Dim uiAppId As LongPtr
    With uiAuto
        Set uiDesktop = .GetRootElement
        Set uiCond_Type = .CreatePropertyCondition(UIA_ControlTypePropertyId, UIA_WindowControlTypeId)
        Set uiCond_Name = .CreatePropertyCondition(UIA_NamePropertyId, "Calculator")
        Set uiCond_And = .CreateAndCondition(uiCond_Type, uiCond_Name)
    End With
    Set uiApp = uiDesktop.FindFirst(TreeScope_Children, uiCond_And)
    If Not uiApp Is Nothing Then
        If uiApp.CurrentIsOffscreen Then
            Set uiAppPattern = uiApp.GetCurrentPattern(UIA_WindowPatternId)
            uiAppPattern.SetWindowVisualState WindowVisualState.WindowVisualState_Normal
            DoEvents
            Sleep 500
        End If
        
        uiAppId = uiApp.GetCurrentPropertyValue(UIA_NativeWindowHandlePropertyId)
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim uiBtn As IUIAutomationElement
    Set uiBtn = uiAuto.ElementFromHandle(ByVal uiAppId)
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim uiBtnCond As IUIAutomationCondition
    Set uiBtnCond = uiAuto.CreatePropertyCondition(UIA_AutomationIdPropertyId, "num3Button")
    
    Dim uiBtnKey As IUIAutomationElement
    Set uiBtnKey = uiBtn.FindFirst(TreeScope_Descendants, uiBtnCond)
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim uiBtnClick As IUIAutomationInvokePattern
    Set uiBtnClick = uiBtnKey.GetCurrentPattern(UIA_InvokePatternId)
    uiBtnClick.Invoke
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim uiResultCond As IUIAutomationCondition
    Set uiResultCond = uiAuto.CreatePropertyCondition(UIA_AutomationIdPropertyId, "CalculatorResults")
    Dim uiResult As IUIAutomationElement
    Set uiResult = uiApp.FindFirst(TreeScope_Descendants, uiResultCond)
    If uiResult Is Nothing Then
        Dim uiResultCond2 As IUIAutomationCondition
        Set uiResultCond2 = uiAuto.CreatePropertyCondition(UIA_AutomationIdPropertyId, "CalculatorAlwaysOnTopResults")
        Dim uiResult2 As IUIAutomationElement
        Set uiResult2 = uiApp.FindFirst(TreeScope_Descendants, uiResultCond2)
        
        MsgBox uiResult2.CurrentName
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim uiExpressionCond As IUIAutomationCondition
    Dim uiExpression As IUIAutomationElement
    Set uiExpressionCond = uiAuto.CreatePropertyCondition(UIA_AutomationIdPropertyId, "CalculatorExpression")
    Set uiExpression = uiApp.FindFirst(TreeScope_Descendants, uiExpressionCond)
    
    If Not uiExpression Is Nothing Then
        MsgBox uiExpression.CurrentName
    End If

    MessageBoxW uiAppId, StrPtr("OK Message"), StrPtr("OK Title"), vbOKOnly

End Sub


