Attribute VB_Name = "sdmx_2_1_Helper"
' ##################################################################
' Author: borisV
' VERSION_DATE: 20200417
' DESCR:
' INFORMATION MODEL:UML CONCEPTUAL DESIGN
' https://sdmx.org/wp-content/uploads/SDMX_2-1-1_SECTION_2_InformationModel_201108.pdf
' http://sdmx.org/wp-content/uploads/SDMX_2-1-1_SECTION_3A_SDMX_ML_201108.zip
' ##################################################################

'SDMX_2_1_SECTION_03A_PART_II_COMMON:

'ObservationDimensionType
Public Type ObservationDimension
    NCNameID As String
    ObsDimensionsCodeType As String
End Type

'com:TextType
Public Type item
    lang As String
    item As String
End Type

'TODO - old delete
Public Type items
    itmName() As sdmx_2_1_Helper.item
    itmDescription() As sdmx_2_1_Helper.item
End Type

'StructurePackageTypeCodelistType
Public Enum StructurePackageTypeCodelist
    datastructure
    metadatastructure
End Enum

'StructureTypeCodelistType
Public Enum StructureTypeCodelist
    datastructure
    metadatastructure
End Enum

'DataStructureRefType
Public Type DataStructureRef
    agencyID As String
    id As String
    local As Boolean
    Class As StructureTypeCodelist
    package As StructurePackageTypeCodelist
End Type

'DataStructureReferenceType
Public Type com_DataStructureReference
'    Ref as dmx_2_1_Helper.DataStructureRef
    URN As String
End Type

' com:PayloadStructureType
' com:StructureSpecificDataTimeSeriesStructureType
' Implemented only StructureSpecificDataTimeSeries elements
Public Type com_Structure
    structureID As String
    schemaURL As String
    namespace As String
    dimensionAtObservation As sdmx_2_1_Helper.ObservationDimension
'    explicitMeasures As Boolean 'PayloadStructureType
    serviceURL As String
    structureURL As String
'    ProvisionAgrement 'PayloadStructureType
'    StructureUsage 'PayloadStructureType
'    Structure 'PayloadStructureType
    agencyID As String
End Type



'SDMX_2_1_SECTION_03A_PART_III_STRUCTURE

'CodeType
Public Type code
    id As String
    URN As String
    uri As String
    'comAnnotations
    comName() As sdmx_2_1_Helper.item
    comDescription() As sdmx_2_1_Helper.item
'    comParent
    item As sdmx_2_1_Helper.items 'TODO remove cCodeList2_1Reader
End Type

'DataflowType
Public Type Dataflow
    id As String
    URN As String
    uri As String
    version As String
    validFrom As Date
    validTo As Date
    agencyID As String
    isFinal As Boolean
    isExternalReference As Boolean
    serviceURL As String
    structureURL As String
    '    comAnnotations
    comName() As sdmx_2_1_Helper.item
    comDescription() As sdmx_2_1_Helper.item
'    Structure() As sdmx_2_1_Helper.com_DataStructureReference
End Type

'CodelistType
Public Type Codelist
    id As String
    URN As String
    uri As String
    version As String
    validFrom As Date
    validTo As Date
    agencyID As String
    isFinal As Boolean
    isExternalReference As Boolean
    serviceURL As String
    structureURL As String
    isPartial As Boolean
'    comAnnotations
    comName() As sdmx_2_1_Helper.item
    comDescription() As sdmx_2_1_Helper.item
    item As sdmx_2_1_Helper.items 'TODO remove CCodeList2_1Reader
    code() As sdmx_2_1_Helper.code
End Type

'DataflowsType
Public Type Dataflows
    Dataflow As String
End Type

'CodelistsType
Public Type Codelists
    Codelist As sdmx_2_1_Helper.Codelist
End Type

'StructuresType
Public Type str_Structure
'    OrganisationSchemes
     Dataflows  As sdmx_2_1_Helper.Dataflows
'    Metadataflows
'    CategorySchemes
'    Categorisations
    Codelist As sdmx_2_1_Helper.Codelists
'    HierarchicalCodelist s
'    Concepts
'    MetadataStructures
'    DataStructures
'    StructureSets
'    ReportingTaxonomies
'    Processes
'    Constraints
'    ProvisionAgreements
End Type


' SDMX_2_1_SECTION_03A_PART_I_MESSAGE:

' BaseHeaderType
' StructureSpecificTimeSeriesDataHeaderType
Public Type Header
    id As String
    test As Boolean
    Prepared As Date
    Sender As String
    Receiver As String
    comName As String
    Structure As sdmx_2_1_Helper.com_Structure
    DataProvider As String
    DataSetAction As String
    DataSetID As String
    Extracted As Date
    ReportingBegin As Date
    ReportingEnd As Date
    EmbargoDate As String
    Source As String
End Type

Public Type DataSet
    todo As Variant
End Type


'StructureSpecificTimeSeriesDataType
Public Type message
    Header As sdmx_2_1_Helper.Header
    DataSet As sdmx_2_1_Helper.DataSet
'    Footer
End Type

'StructureType
Public Type messageStructures
    Header As sdmx_2_1_Helper.Header
'    Structures As 'str:StructuresType
'    Footer
End Type
