'find duplicate records in ArcMap tables
'Can be shapefile or geodatabase
'modified from VBA code downloaded from ESRI forum
'converted to V.Net ESRI Add-in for ArcMap 10+
'original writer unknown
'Conversion: Grant Herbert 2010
'Hosted on GitHub 2013
'updated 2013 using version 10.1 SDK, recompile with older SDK if required for version 10

Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geometry
Imports System.Windows.Forms

Public Class frmFindDups
    Dim pMxDoc As IMxDocument
    Private m_pMaps As Maps
    Private m_pMap As IMap
    Public Sub New(ByRef M_hook As ESRI.ArcGIS.Framework.IDocument)
        InitializeComponent()
        'hook passed through from button on_click event
        pMxDoc = M_hook
        m_pMaps = pMxDoc.Maps
        m_pMap = pMxDoc.FocusMap

    End Sub

    Public Enum TableType
        NoType = 0
        IStandaloneTable = 1
        IFeatureLayer = 2
    End Enum
    Private m_pTable As ITable
    Private m_TabType As TableType

    ' frmFindDups - Duplicate Record Tagger
    ' Selects duplicate records in selected table

    Private Sub btnGetFields_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetFields.Click

        Dim pSelItem As Object
        Dim TabType As TableType
        Dim pTable As ITable

        ' Check for selected table or feature class

        m_pTable = Nothing
        m_TabType = TableType.NoType
        lbFields.Items.Clear()
        cbVisExtent.Checked = False

        pSelItem = pMxDoc.SelectedItem
        If pSelItem Is Nothing Then 'no table selected in TOC
            MsgBox("Nothing selected", vbExclamation)
            Exit Sub
        ElseIf TypeOf pSelItem Is IStandaloneTable Then
            TabType = TableType.IStandaloneTable
        ElseIf TypeOf pSelItem Is IFeatureLayer Then
            TabType = TableType.IFeatureLayer
        Else
            MsgBox("Unable to work with this object", vbExclamation)
            Exit Sub
        End If
        pTable = pSelItem

        ' Populate field list from selected table

        Dim pFields As ITableFields
        Dim i As Long
        Dim pInfo As IFieldInfo
        Dim pField As IField
        Dim fType As esriFieldType

        pFields = pTable
        For i = 0 To pFields.FieldCount - 1
            pInfo = pFields.FieldInfo(i)
            If pInfo.Visible Then

                ' Check for supported field type

                pField = pTable.Fields.Field(i)
                fType = pField.Type
                If Not (fType = esriFieldType.esriFieldTypeBlob Or _
                   fType = esriFieldType.esriFieldTypeGeometry Or _
                   fType = esriFieldType.esriFieldTypeRaster) Then

                    lbFields.Items.Add(pField.AliasName)

                End If
            End If
        Next
        If lbFields.Items.Count = 0 Then
            MsgBox("No valid fields", vbExclamation)
            Exit Sub
        End If

        If TabType = TableType.IStandaloneTable Then
            cbVisExtent.Enabled = False
        Else
            cbVisExtent.Enabled = True
        End If

        m_pTable = pTable
        m_TabType = TabType
        ToolStripStatusLabel1.Text = "Fields Loaded"
        Application.DoEvents()
    End Sub

    Private Sub btnFindDups_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindDups.Click

        If m_pTable Is Nothing Then
            MsgBox("No table selected", vbExclamation)
            Exit Sub
        End If

        Dim pMouse As IMouseCursor
        pMouse = New MouseCursor
        pMouse.SetCursor(2)

        ' Get selected field(s)

        Dim FList() As String = Nothing, FieldString As String, FName As String
        Dim i As Long, n As Long, iField As Long

        FieldString = ""
        FName = ""
        n = 0

        If lbFields.SelectedItem <> "" Then 'check if a field selected
            Debug.Print(lbFields.SelectedItem.ToString)
            n = lbFields.SelectedIndex 'get index number in list
            Debug.Print(n)
            iField = m_pTable.Fields.FindFieldByAliasName(lbFields.SelectedItem)
            Debug.Print(iField)
            FName = m_pTable.Fields.Field(iField).Name
            Debug.Print(FName)
            ReDim Preserve FList(n)
            FList(n) = FName
            If FieldString = "" Then
                FieldString = FName
            Else
                FieldString = FieldString & ", " & FName
            End If
            'n = n + 1

        End If
        'Next
        If n = 0 Then
            MsgBox("No field(s) selected", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        ToolStripStatusLabel1.Text = "Field: " & FName
        Application.DoEvents()
        ' Check for selection set

        Dim bUseSel As Boolean
        Dim pTSel As ITableSelection = Nothing
        Dim pFSel As IFeatureSelection = Nothing
        Dim pSS As ISelectionSet

        bUseSel = False
        If m_TabType = TableType.IStandaloneTable Then
            pTSel = m_pTable
            pSS = pTSel.SelectionSet
        ElseIf m_TabType = TableType.IFeatureLayer Then
            pFSel = m_pTable
            pSS = pFSel.SelectionSet
        Else
            MsgBox("Not a Table or FeatureLayer selection", MsgBoxStyle.OkOnly)
            Exit Sub
        End If
        i = pSS.Count
        bUseSel = (i > 0)

        ' Check for definition query in the table

        Dim bUseQF As Boolean
        Dim pTD As ITableDefinition
        Dim WhereClause As String
        Dim pQF As IQueryFilter = Nothing

        WhereClause = ""

        bUseQF = False
        If (Not bUseSel) Then
            pTD = m_pTable
            WhereClause = pTD.DefinitionExpression
            If WhereClause <> "" Then
                pQF = New QueryFilter
                pQF.WhereClause = WhereClause
                bUseQF = True
            End If
        End If

        ' Check for restrict to visible extent setting

        Dim bUseSF As Boolean
        Dim pActiveView As IActiveView
        Dim pViewExtent As IEnvelope
        Dim pPC As IPointCollection
        Dim pPoly As IPolygon
        Dim pSF As ISpatialFilter = Nothing

        bUseSF = False
        If cbVisExtent.Checked Then
            'pMxDoc = ThisDocument
            'pMap = pMxDoc.FocusMap
            pActiveView = m_pMap
            pViewExtent = pActiveView.Extent
            pPC = New Polygon
            With pPC
                .AddPoint(pViewExtent.LowerLeft)
                .AddPoint(pViewExtent.UpperLeft)
                .AddPoint(pViewExtent.UpperRight)
                .AddPoint(pViewExtent.LowerRight)
            End With
            pPoly = pPC
            pPoly.Close()
            pPoly.SpatialReference = m_pMap.SpatialReference
            pSF = New SpatialFilter
            With pSF
                .Geometry = pPoly
                .GeometryField = "SHAPE"
                .SpatialRel = esriSpatialRelEnum.esriSpatialRelContains
                If bUseQF Then
                    .WhereClause = WhereClause
                End If
            End With
            bUseSF = True
        End If

        ' Sort table

        ToolStripStatusLabel1.Text = "Sorting..."
        Application.DoEvents()
        Dim pTS As ITableSort
        pTS = New TableSort
        With pTS
            .Fields = FieldString
            .Table = m_pTable
            For i = 0 To n - 1
                .Ascending(FList(i)) = True
                If cbCaseSensitive.Checked Then
                    .CaseSensitive(FList(i)) = True
                End If
            Next
            If bUseSF Then
                .QueryFilter = pSF
            ElseIf bUseQF Then
                .QueryFilter = pQF
            End If
            If bUseSel Then
                .SelectionSet = pSS
            End If
        End With
        pTS.Sort(Nothing)

        ' Loop through rows and find duplicates

        Dim pCursor As ICursor
        Dim pRow1 As IRow, pRow2 As IRow
        Dim pFeature As IFeature
        Dim bSelectAll As Boolean, bDup As Boolean, bFirst As Boolean
        Dim iNumRecs As Long, iNumDups As Long

        ToolStripStatusLabel1.Text = "Analyzing..."
        Application.DoEvents()
        bSelectAll = cbSelectAll.Checked
        pCursor = pTS.Rows
        pRow1 = pCursor.NextRow
        pRow2 = pCursor.NextRow
        iNumRecs = 0
        iNumDups = 0
        If m_TabType = TableType.IStandaloneTable Then
            pTSel.Clear()
        ElseIf m_TabType = TableType.IFeatureLayer Then
            pFSel.Clear()
        End If
        bFirst = True
        Do While Not pRow2 Is Nothing

            iNumRecs = iNumRecs + 1
            If iNumRecs Mod 100 = 0 Then
                StatusStrip1.Text = "Analyzing... " & iNumRecs
                Application.DoEvents()
            End If

            bDup = True
            For i = 0 To n - 1
                iField = m_pTable.FindField(lbFields.SelectedItem)
                Debug.Print(pRow1.Value(iField) & " , " & pRow2.Value(iField))
                If pRow1.Value(iField) Is DBNull.Value Or pRow2.Value(iField) Is DBNull.Value Then 'check for nulls - skip
                    bDup = False
                    Exit For
                End If
                If pRow1.Value(iField) <> pRow2.Value(iField) Then 'problem with DBNull here, hence the skip earlier
                    bDup = False
                    Exit For
                End If
            Next
            If bDup Then
                If m_TabType = TableType.IStandaloneTable Then
                    If bFirst And bSelectAll Then
                        pTSel.AddRow(pRow1)
                    End If
                    pTSel.AddRow(pRow2)
                ElseIf m_TabType = TableType.IFeatureLayer Then
                    If bFirst And bSelectAll Then
                        pFeature = pRow1
                        pFSel.Add(pFeature)
                    End If
                    pFeature = pRow2
                    pFSel.Add(pFeature)
                End If
                iNumDups = iNumDups + 1
                bFirst = False
            Else
                bFirst = True
            End If
            pRow1 = pRow2
            pRow2 = pCursor.NextRow

        Loop

        ' Wrap up

        If m_TabType = TableType.IStandaloneTable Then
            pTSel.SelectionChanged()
            pSS = pTSel.SelectionSet
        ElseIf m_TabType = TableType.IFeatureLayer Then
            pFSel.SelectionChanged()
            pSS = pFSel.SelectionSet
        End If
        i = pSS.Count
        ToolStripStatusLabel1.Text = "Completed - Found " & i & " Duplicates"
        Application.DoEvents()
        MsgBox("Duplicates found: " & i, MsgBoxStyle.OkOnly)

    End Sub



    Private Sub btnHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHelp.Click
        MsgBox("Selects duplicate records in an ArcMap table or feature layer based on selected field." & Environment.NewLine & _
               "The first record is left unselected by default. Manual verification is suggested." & Environment.NewLine & _
               "Works on the entire table or a selection of records." & Environment.NewLine & _
               "1. Select layer or table in the TOC and click 'Get Field List'." & Environment.NewLine & _
               "2. Select criteria field and click 'Find Duplicates'." & Environment.NewLine & _
               "3. The 'Select All' option selects ALL matching records instead of just the duplicates." & Environment.NewLine & _
               "4. The 'Restrict to Visible Extent' option restricts feature class results to the active view." & Environment.NewLine & _
               "5. The 'Reset' button clears selection & the field list." & Environment.NewLine & _
               "-> a RESET is recommended between tables." & Environment.NewLine & _
               "Version for ESRI 10.1, 2013", MsgBoxStyle.OkOnly, "Find Duplicates Help")
    End Sub

    Private Sub btnRestart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRestart.Click
        'Restart - clear the selection in the table and clear the field list
        lbFields.Items.Clear()
        m_pMap.ClearSelection() 'clears all selections in the map
        ToolStripStatusLabel1.Text = "Ready"
        Application.DoEvents()
    End Sub

    Private Sub frmFindDups_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

End Class