''' Option Explicit
!INC Local Scripts.EAConstants-VBScript

''' ===========================================================================
''' TAGGEDVALUE CONSTANTS
''' ===========================================================================
'''
''' VERSION			: 0.9.7, 20200928 - Updated header text.
'''  				: 0.9.6, 20151210 - Initial commit.

Public Const EA_ASSOCIATION_SOURCE = "ASSOCIATION_SOURCE"
Public Const EA_ASSOCIATION_TARGET = "ASSOCIATION_TARGET"

Public Const EA_UNSPECIFIED			= "Unspecified"
Public Const EA_BIDIRECTIONAL		= "Bi-Directional"
Public Const EA_SOURCEDESTINATION	= "Source -> Destination"
Public Const EA_DESTINATIONSOURCE	= "Destination -> Source"

Public Const EA_none			= 0		''' "none"
Public Const EA_shared			= 1		''' "shared"
Public Const EA_composite		= 2		''' "composite"


''' MODEL ELEMENTS

Private Const EA_Element 		= 4		''' Class & Interface (see <stereotype>!)
Private Const EA_Class 			= 4		''' -"-
Private Const EA_Interface		= 4		''' -"-
Private Const EA_Package		= 5		''' Package
Private Const EA_Attribute		= 23	''' Attributes
Private Const EA_Method			= 24	''' Methods
Private Const EA_Connector		= 7		''' Connectors
Private Const EA_Role			= 22	''' Role/ConnectorEnd
Private Const EA_ConnectorEnd	= 22	''' Role/ConnectorEnd

''' -----------------------------------------------------
''' ELEMENT TYPES
''' -----------------------------------------------------

Const METATYPE_CLASS = "Class"
Const METATYPE_INTERFACE = "Interface"
Const METATYPE_GENERALIZATION = "Generalization"
Const METATYPE_REALISATION  = "Realisation"
Const METATYPE_DEPENDENCY  = "Dependency"
Const METATYPE_ASSOCIATION  = "Association" ' + Aggregation  == Association
Const METATYPE_AGGREGATION  = "Aggregation"
Const METATYPE_COMPOSITION = "Composition"

Private Const msg_NilPointer 		= "Access Violation: Nil Pointer!"
Private Const msg_NilPointer_Subject = "Nil Pointer Error! Subject Class has not been assigned."
Private Const msg_UnknownSubjectID 	= "Fatal error: SubjectID is Unknown. Doesn't match ClientID or SupplierID!"
Private Const msg_AggregationKind	= "Invalid AggregationKind in WConnectionEnd! (This is a Bug)"

Dim err_NilPointer 			: err_NilPointer 		= vbObjectError + 1100
Dim err_UnknownSubjectID 	: err_UnknownSubjectID	= vbObjectError + 1101
Dim err_AggregationKind 	: err_AggregationKind	= vbObjectError + 1102
