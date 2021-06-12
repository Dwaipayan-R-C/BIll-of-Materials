
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports System.IO
Module Module1
    Structure BomPosition
        Dim ModelPath As String
        Dim Configuration As String
        Dim Quantity As Double
        Dim Description As String
        Dim Price As Double
    End Structure
    Sub Main()
        Dim appXL As New Excel.Application
        Dim wbXl As Excel.Workbook
        Dim shXL As Excel.Worksheet
        Dim raXL As Excel.Range

        ' Start Excel and get Application object.
        appXL = CreateObject("Excel.Application")

        ' Add a new workbook.
        wbXl = appXL.Workbooks.Add
        shXL = wbXl.ActiveSheet

        ' Add table headers going cell by cell.
        shXL.Cells(1, 1).Value = "Sl No."
        shXL.Cells(1, 2).Value = "Path"
        shXL.Cells(1, 3).Value = "Configuration"
        shXL.Cells(1, 4).Value = "Description"
        shXL.Cells(1, 5).Value = "Quantity"
        shXL.Cells(1, 6).Value = "Price"

        shXL.Range("B2").ColumnWidth = 100
        shXL.Range("A2").ColumnWidth = 8
        shXL.Range("C2").ColumnWidth = 20
        shXL.Range("D2").ColumnWidth = 40
        shXL.Range("E2").ColumnWidth = 10
        shXL.Range("F2").ColumnWidth = 10
        ' Format A1:D1 as bold, vertical alignment = center.
        With shXL.Range("A1", "F1")
            .Font.Bold = True
            .Font.Size = 14
            .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            .HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter
        End With

        Dim swApp As SldWorks.SldWorks = CreateObject("SldWorks.Application")
        Dim swModelDoc As SldWorks.ModelDoc2 = swApp.OpenDoc6("C:\GeneratedModel.SLDASM", SwConst.swDocumentTypes_e.swDocASSEMBLY, SwConst.swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0)
        swModelDoc.Visible = True
        Dim swAssy As SldWorks.AssemblyDoc = swModelDoc

        If Not swAssy Is Nothing Then
            swAssy.ResolveAllLightWeightComponents(True)
            Dim bom() As BomPosition
            bom = GetFlatBom(swAssy)
            Dim i As Integer
            For i = 0 To UBound(bom)
                shXL.Cells(i + 2, 1).Value = i + 1
                shXL.Cells(i + 2, 2).Value = bom(i).ModelPath
                shXL.Cells(i + 2, 3).Value = bom(i).Configuration
                shXL.Cells(i + 2, 4).Value = bom(i).Description
                shXL.Cells(i + 2, 5).Value = bom(i).Quantity
                shXL.Cells(i + 2, 6).Value = bom(i).Price
                Debug.Print(bom(i).ModelPath & vbTab & bom(i).Configuration & vbTab & bom(i).Description & vbTab & bom(i).Price & vbTab & bom(i).Quantity)
                shXL.Range(shXL.Cells(i + 2, 1), shXL.Cells(i + 2, 6)).BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic)
            Next

            shXL.Range("A1", "A" + (i + 2).ToString).BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic)
            shXL.Range("B1", "B" + (i + 2).ToString).BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic)
            shXL.Range("C1", "C" + (i + 2).ToString).BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic)
            shXL.Range("D1", "D" + (i + 2).ToString).BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic)
            shXL.Range("E1", "E" + (i + 2).ToString).BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic)
            shXL.Range("F1", "F" + (i + 2).ToString).BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic)
            With shXL.Range("A2", "F" + (i + 2).ToString)
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                .HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
        Else
            MsgBox("Please open assembly")
        End If
        ' Make sure Excel is visible and give the user control
        ' of Excel's lifetime.
        appXL.UserControl = True
        appXL.Visible = True
        ' Release object references.
        raXL = Nothing
        shXL = Nothing
        wbXl = Nothing

    End Sub
    Function GetFlatBom(assy As SldWorks.AssemblyDoc) As BomPosition()
        Dim bom() As BomPosition
        Dim isInit As Boolean
        Dim vComps As Object
        vComps = assy.GetComponents(False)
        Dim i As Integer
        For i = 0 To UBound(vComps)
            Dim swComp As SldWorks.Component2
            swComp = vComps(i)
            If swComp.GetSuppression() <> 0 And Not swComp.ExcludeFromBOM Then
                Dim bomPos As Integer
                bomPos = FindBomPosition(bom, swComp)
                If bomPos = -1 Then
                    If isInit Then
                        ReDim Preserve bom(UBound(bom) + 1)
                    Else
                        isInit = True
                        ReDim bom(0)
                    End If
                    bomPos = UBound(bom)
                    bom(bomPos).ModelPath = swComp.GetPathName()
                    bom(bomPos).Configuration = swComp.ReferencedConfiguration
                    bom(bomPos).Quantity = 1
                    GetProperties(swComp, bom(bomPos).Description, bom(bomPos).Price)
                Else
                    bom(bomPos).Quantity = bom(bomPos).Quantity + 1
                End If
            End If
        Next
        GetFlatBom = bom

    End Function
    Function FindBomPosition(bom() As BomPosition, comp As SldWorks.Component2) As Integer
        FindBomPosition = -1
        Try
            Dim i As Integer
            For i = 0 To UBound(bom)
                If LCase(bom(i).ModelPath) = LCase(comp.GetPathName()) And LCase(bom(i).Configuration) = LCase(comp.ReferencedConfiguration) Then
                    FindBomPosition = i
                    Exit Function
                End If
            Next
            Exit Try
        Catch
            Exit Try
        End Try
        Return FindBomPosition
    End Function

    Function GetProperties(comp As SldWorks.Component2, ByRef desc As String, ByRef prc As Double) As Object
        Dim swCompModel As SldWorks.ModelDoc2
        swCompModel = comp.GetModelDoc2()
        If swCompModel Is Nothing Then
            Err.Raise("", "Failed to get model from the component")
        End If
        desc = GetPropertyValue(swCompModel, comp.ReferencedConfiguration, "Description")
        On Error Resume Next
        prc = GetPropertyValue(swCompModel, comp.ReferencedConfiguration, "Price")
    End Function
    Function GetPropertyValue(model As SldWorks.ModelDoc2, conf As String, prpName As String) As String
        Dim confSpecPrpMgr As SldWorks.CustomPropertyManager
        Dim genPrpMgr As SldWorks.CustomPropertyManager
        confSpecPrpMgr = model.Extension.CustomPropertyManager(conf)
        genPrpMgr = model.Extension.CustomPropertyManager("")
        Dim prpVal As String
        Dim prpResVal As String
        confSpecPrpMgr.Get3(prpName, False, "", prpVal)
        If prpVal = "" Then
            genPrpMgr.Get3(prpName, False, prpVal, prpResVal)
        End If
        GetPropertyValue = prpResVal
    End Function
End Module
