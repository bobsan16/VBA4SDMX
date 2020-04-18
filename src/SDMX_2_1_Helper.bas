Attribute VB_Name = "sdmx_2_1_Helper"
' ##################################################################
' Author: BorisV
' VERSION_DATE: 20200418
' DESCR:
' INFORMATION MODEL:UML CONCEPTUAL DESIGN
' https://sdmx.org/wp-content/uploads/SDMX_2-1-1_SECTION_2_InformationModel_201108.pdf
' http://sdmx.org/wp-content/uploads/SDMX_2-1-1_SECTION_3A_SDMX_ML_201108.zip
' ##################################################################


'SDMX_2_1_SECTION_03A_PART_II_COMMON:


'TimeDataType
Public Enum com_TimeData
    ObservationalTimePeriod
    StandardTimePeriod
    BasicTimePeriod
    GregorianTimePeriod
    GregorianYear
    GregorianYearMonth
    GregorianDay
    ReportingTimePeriod
    ReportingYear
    ReportingSemester
    ReportingTrimester
    ReportingQuarter
    ReportingMonth
    ReportingWeek
    ReportingDay
    DateTime
    TimeRange
End Enum

'ItemTypeCodelistType
Public Enum com_ItemTypeCodelist
    Agency
    DataConsumer
    DataProvider
    OrganisationUnit
End Enum

'ItemSchemeTypeCodelistType
Public Enum com_ItemSchemeTypeCodelist
    AgencyScheme
    categoryscheme
    Codelist
    ConceptScheme
    DataConsumerScheme
    DataProviderScheme
    OrganisationUnitScheme
    ReportingTaxonomy
End Enum

'ItemSchemePackageTypeCodelistType
Public Enum com_ItemSchemePackageTypeCodelist
    Base
    Codelist
    categoryscheme
    ConceptScheme
End Enum

'DimensionTypeType
Public Enum com_DimensionType
    Dimension
    MeasureDimension
    TimeDimension
End Enum

'StructurePackageTypeCodelistType
Public Enum com_StructurePackageTypeCodelist
    DataStructure
    metadatastructure
End Enum

'StructureTypeCodelistType
Public Enum com_StructureTypeCodelist
    DataStructure
    metadatastructure
End Enum

'ObservationDimensionType
Public Type com_ObservationDimension
    NCNameID As String
    ObsDimensionsCodeType As String
End Type

'com:TextType
Public Type com_item
    lang As String
    item As String
End Type

'TODO - old delete
Public Type com_items
    itmName() As sdmx_2_1_Helper.com_item
    itmDescription() As sdmx_2_1_Helper.com_item
End Type

'LocalDimensionRefType
Public Type com_LocalDimensionRef
    id As String
    local As Boolean 'fixed: true
'    Class As 'default: Dimension
    package As com_StructurePackageTypeCodelist 'fixed: datastructure
End Type

'LocalDimensionReferenceType
Public Type com_LocalDimensionReference
    Ref As sdmx_2_1_Helper.com_LocalDimensionRef
End Type

'ConceptRefType
Public Type com_ConceptRef
    agencyID As String
'    maintainableParentID
'    maintainableParentVersion
    id As String
    local As Boolean
    Class As com_ItemSchemeTypeCodelist 'fixed: Codelist
    package As com_ItemSchemePackageTypeCodelist 'fixed: codelist
End Type

'ConceptReferenceType
Public Type com_ConceptReference
    Ref As sdmx_2_1_Helper.com_ConceptRef
    urn() As String
End Type

'CodelistRefType
Public Type com_CodelistRef
    agencyID As String
    id As String
    version  As String
    local As Boolean
    Class As com_ItemSchemeTypeCodelist 'fixed: Codelist
    package As com_ItemSchemePackageTypeCodelist 'fixed: codelist
End Type

'CodelistReferenceType
Public Type com_CodelistReference
    Ref As sdmx_2_1_Helper.com_CodelistRef
    urn() As String
End Type


'DataStructureRefType
Public Type com_DataStructureRef
    agencyID As String
    id As String
    local As Boolean
    Class As com_StructureTypeCodelist
    package As com_StructurePackageTypeCodelist
End Type

'DataStructureReferenceType
Public Type com_DataStructureReference
    Ref As sdmx_2_1_Helper.com_DataStructureRef
    urn() As String
End Type

' com:PayloadStructureType
' com:StructureSpecificDataTimeSeriesStructureType
' Implemented only StructureSpecificDataTimeSeries elements
Public Type com_Structure
    structureID As String
    schemaURL As String
    namespace As String
    dimensionAtObservation As sdmx_2_1_Helper.com_ObservationDimension
'    explicitMeasures As Boolean 'PayloadStructureType
    serviceURL As String
    structureURL As String
'    ProvisionAgrement 'PayloadStructureType
'    StructureUsage 'PayloadStructureType
'    Structure 'PayloadStructureType
    agencyID As String
End Type



'SDMX_2_1_SECTION_03A_PART_III_STRUCTURE

Public Enum str_UsageStatus
    Mandatory
    Conditional
End Enum


'CodeType
Public Type str_code
    id As String
    urn As String
    uri As String
    'comAnnotations
    comName() As sdmx_2_1_Helper.com_item
    comDescription() As sdmx_2_1_Helper.com_item
'    comParent
    item As sdmx_2_1_Helper.com_items 'TODO remove cCodeList2_1Reader
End Type

'CodelistType
Public Type str_Codelist
    id As String
    urn As String
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
    comName() As sdmx_2_1_Helper.com_item
    comDescription() As sdmx_2_1_Helper.com_item
    item As sdmx_2_1_Helper.com_items 'TODO remove CCodeList2_1Reader
    code() As sdmx_2_1_Helper.str_code
End Type

'DataflowType
Public Type str_Dataflow
    id As String
    urn As String
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
    comName() As sdmx_2_1_Helper.com_item
    comDescription() As sdmx_2_1_Helper.com_item
    Structure() As sdmx_2_1_Helper.com_DataStructureReference
End Type

'ContactType
Public Type str_Contact
    id As String
    comName() As sdmx_2_1_Helper.com_item
    Department As sdmx_2_1_Helper.com_item
    Role As sdmx_2_1_Helper.com_item
    Telephone() As String
    Fax() As String
    X400() As String
    uri() As String
    Email() As String
End Type

'ConceptType
Public Type str_Concept
     id As String
     urn As String
     uri As String
'     comAnnotations
    comName() As sdmx_2_1_Helper.com_item
    comDescription() As sdmx_2_1_Helper.com_item
'    Parent As
'    CoreRepresentation
'    ISOConceptReference
End Type

'AgencyType
Public Type str_Agency
     id As String
     urn As String
     uri As String
'     comAnnotations
    comName() As sdmx_2_1_Helper.com_item
    comDescription() As sdmx_2_1_Helper.com_item
    Contact As sdmx_2_1_Helper.str_Contact
End Type

'DataflowsType
Public Type str_Dataflows
    Dataflow As String
End Type

'CodelistsType
Public Type str_Codelists
    Codelist As sdmx_2_1_Helper.str_Codelist
End Type

'AgencySchemeType
Public Type str_AgencyScheme
    id As String '(fixed: AGENCIES)
    urn As String
    uri As String
    version As String '(fixed: 1.0)
    validFrom As Date
    validTo As Date
    agencyID As String
    isFinal As Boolean '(fixed: false)
    isExternalReference As Boolean '(default: false)
    serviceURL As String
    structureURL As String
    isPartial As Boolean '(default: false)
'    comAnnotations
    comName() As sdmx_2_1_Helper.com_item
    comDescription() As sdmx_2_1_Helper.com_item
    Agency As sdmx_2_1_Helper.str_Agency
End Type

'OrganisationSchemesType
Public Type str_OrganisationSchemes
    AgencyScheme As sdmx_2_1_Helper.str_AgencyScheme
'    DataConsumerScheme
'    DataProviderScheme
'    OrganisationUnitScheme
End Type


'ConceptSchemeType
Public Type str_ConceptScheme
    id As String
    urn As String
    uri As String
    version As String '(default: 1.0)
    validFrom As Date
    validTo As Date
    agencyID As String
    isFinal As Boolean '(default: false)
    isExternalReference As Boolean '(default: false)
    serviceURL As String
    structureURL As String
    isPartial As Boolean '(default: false)
'    comAnnotations
    comName() As sdmx_2_1_Helper.com_item
    comDescription() As sdmx_2_1_Helper.com_item
    Concept() As sdmx_2_1_Helper.str_Concept
End Type

'ConceptsType
Public Type str_Concepts
    ConceptScheme() As sdmx_2_1_Helper.str_ConceptScheme
End Type

'SimpleDataStructureRepresentationType
Public Type str_SimpleDataStructureRepresentation
'    TextFormat
    enumeration As sdmx_2_1_Helper.com_CodelistReference
'    EnumerationFormat
End Type

'DimensionType
Public Type str_Dimension
    id As String
    urn As String
    uri As String
    position As Long
    type As com_DimensionType '(fixed: Dimension)
'    comAnnotations
    comConceptIdentity As sdmx_2_1_Helper.com_ConceptReference
    LocalRepresentation As sdmx_2_1_Helper.str_SimpleDataStructureRepresentation
'    comConceptRole
End Type

'TimeTextFormatType
Public Type str_TimeTextFormat
    textType As com_TimeData 'default: ObservationalTimePeriod
    startTime As Date
    endTime As Date
End Type

'TimeDimensionRepresentationType
Public Type str_TimeDimensionRepresentation
    TextFormat As sdmx_2_1_Helper.str_TimeTextFormat
End Type

'TimeDimensionType
Public Type str_TimeDimension
    id As String 'fixed: TIME_PERIOD
    urn As String
    uri As String
    position As Long
'    type as 'fixed: TimeDimension
'    comAnnotations
    comConceptIdentity As sdmx_2_1_Helper.com_ConceptReference
    LocalRepresentation As sdmx_2_1_Helper.str_TimeDimensionRepresentation
End Type

'DimensionListType
Public Type str_DimensionList
    id As String 'fixed: DimensionDescriptor
    urn As String
    uri As String
'    comAnnotations
    Dimension() As sdmx_2_1_Helper.str_Dimension
'    MeasureDimension
    TimeDimension As sdmx_2_1_Helper.str_TimeDimension
End Type

'AttributeRelationshipType
Public Type str_AttributeRelationship
    None As Variant
    Dimension As sdmx_2_1_Helper.com_LocalDimensionReference
'    AttachmentGroup
'    Group
'    PrimaryMeasure
End Type
'AttributeType
Public Type str_Attribute
    id As String
    urn As String
    uri As String
    assignmentStatus As str_UsageStatus
'    comAnnotations
    comConceptIdentity As sdmx_2_1_Helper.com_ConceptReference
    LocalRepresentation As sdmx_2_1_Helper.str_SimpleDataStructureRepresentation
'    comConceptRole
    AttributeRelationship As sdmx_2_1_Helper.str_AttributeRelationship
End Type

'AttributeListType
Public Type str_AttributeList
    id As String 'fixed: AttributeDescriptor
    urn As String
    uri As String
'    comAnnotations
    Attribute As sdmx_2_1_Helper.str_Attribute
'    ReportingYearStartDay
End Type

'PrimaryMeasureType
Public Type str_PrimaryMeasure
    id As String 'fixed: OBS_VALUE
    urn As String
    uri As String
'    comAnnotations
    comConceptIdentity As sdmx_2_1_Helper.com_ConceptReference
    LocalRepresentation As sdmx_2_1_Helper.str_SimpleDataStructureRepresentation
End Type

'MeasureListType
Public Type str_MeasureList
    id As String 'fixed: MeasureDescriptor
    urn As String
    uri As String
'    comAnnotations
    PrimaryMeasure As sdmx_2_1_Helper.str_PrimaryMeasure
End Type

'DataStructureComponentsType
Public Type str_DataStructureComponents
    DimensionList As sdmx_2_1_Helper.str_DimensionList
'    Group
    AttributeList As sdmx_2_1_Helper.str_AttributeList
    MeasureList As sdmx_2_1_Helper.str_MeasureList
End Type

'DataStructureType
Public Type str_DataStructure
    id As String
    urn As String
    uri As String
    version As String '(default: 1.0)
    validFrom As Date
    validTo As Date
    agencyID As String
    isFinal As Boolean '(default: false)
    isExternalReference As Boolean '(default: false)
    serviceURL As String
    structureURL As String
'    comAnnotations
    comName() As sdmx_2_1_Helper.com_item
    comDescription() As sdmx_2_1_Helper.com_item
    DataStructureComponents() As sdmx_2_1_Helper.str_DataStructureComponents
End Type

'DataStructuresType
Public Type str_DataStructures
    DataStructure() As sdmx_2_1_Helper.str_DataStructure
End Type

'StructuresType
Public Type str_Structure
    OrganisationSchemes As sdmx_2_1_Helper.str_OrganisationSchemes
    Dataflows  As sdmx_2_1_Helper.str_Dataflows
'    Metadataflows
'    CategorySchemes
'    Categorisations
    Codelist As sdmx_2_1_Helper.str_Codelists
'    HierarchicalCodelist s
    Concepts As sdmx_2_1_Helper.str_Concepts
'    MetadataStructures
    DataStructures As sdmx_2_1_Helper.str_DataStructures
'    StructureSets
'    ReportingTaxonomies
'    Processes
'    Constraints
'    ProvisionAgreements
End Type


' SDMX_2_1_SECTION_03A_PART_I_MESSAGE:

' BaseHeaderType
' StructureSpecificTimeSeriesDataHeaderType
Public Type mes_Header
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

Public Type mes_DataSet
    todo As Variant
End Type

'StructureSpecificTimeSeriesDataType
Public Type mes_Message
    Header As sdmx_2_1_Helper.mes_Header
    DataSet As sdmx_2_1_Helper.mes_DataSet
'    Footer
End Type

'StructureType
Public Type mes_MessageStructures
    Header As sdmx_2_1_Helper.mes_Header
'    Structures As 'str:StructuresType
'    Footer
End Type
