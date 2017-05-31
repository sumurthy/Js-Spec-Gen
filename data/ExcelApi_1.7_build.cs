// --------------------------------------------------------------------------------------------------
// 
// <copyright file="ExcelApi.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>
// <summary>
// Contains the metadata of Excel API that is currently implemented.
// The following is the workflow to add a new API
// 1) DEV add the API to xlshared\src\api\metadata\current\ExcelApi.cs
// 2) DEV runs xlshared\util\XlsApiGen.bat to re-generate the following files
//      xlshared\src\api\Xlapi.h                COM CoClass header file
//      xlshared\src\api\Xlapi_i.h              COM interface header file
//      xlshared\src\api\Xlapi_i.cpp            COM GUIDs
//      xlshared\src\api\TypeRegistration.cpp   Type registration file
//      xlshared\src\api\*.disp.cpp             COM IDispatch interface related implementation
//      xlshared\src\api\script\Xlapi.ts        TypeScript file
// 3) DEV implement the new API, update xlshared\src\api\sources.inc if necessary.
// </summary>
// --------------------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using Microsoft.OfficeExtension.CodeGen;

[assembly: ClientCallableNamespaceMap("Microsoft.ExcelServices", ComCoClassNamespaceName = "ExcelApiImpl", ComInterfaceNamespaceName = "ExcelApi", TypeScriptNamespaceName = "Excel")]

// Default error (fallback if not uniquely mapped below)
[assembly: HResultDefaultError(HttpStatusCode.InternalServerError, Microsoft.ExcelServices.ErrorCodes.GeneralException, "stridsApiGeneralException")]

// Errors we specifically want to hide into general exception (500)
[assembly: HResultError("hrFail", HttpStatusCode.InternalServerError, Microsoft.ExcelServices.ErrorCodes.GeneralException, "stridsApiGeneralException")]
[assembly: HResultError("hrUnexpected", HttpStatusCode.InternalServerError, Microsoft.ExcelServices.ErrorCodes.GeneralException, "stridsApiGeneralException")]
[assembly: HResultError("hrOutOfMemory", HttpStatusCode.InternalServerError, Microsoft.ExcelServices.ErrorCodes.GeneralException, "stridsApiGeneralException")]
[assembly: HResultError("SharedInterimIfs::hrFormulaParseError", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.InvalidArgument, "stridsApiInvalidArgument")]

// Errors 400s
[assembly: HResultError("E_POINTER", HttpStatusCode.NotFound, Microsoft.ExcelServices.ErrorCodes.ItemNotFound, "stridsApiItemNotFound")]
[assembly: HResultError("hrBadIndex", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.InvalidArgument, "stridsWCOUTOFBOUNDS")]
[assembly: HResultError("hrInvalidArg", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.InvalidArgument, "stridsApiInvalidArgument")]
[assembly: HResultError("hrInvalidAPIOperation", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.InvalidOperation, "stridsApiInvalidAPIOperation")]
[assembly: HResultError("hrInvalidBinding", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.InvalidBinding, "stridsApiInvalidBinding")]
[assembly: HResultError("hrInvalidAPISelection", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.InvalidSelection, "stridsApiInvalidSelection")]
[assembly: HResultError("hrInvalidAPIReference", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.InvalidReference, "stridsApiInvalidReference")]
[assembly: HResultError("hrNotFound", HttpStatusCode.NotFound, Microsoft.ExcelServices.ErrorCodes.ItemNotFound, "stridsApiItemNotFound")]
[assembly: HResultError("SharedInterimIfs::hrInsDelDisallowedByFeature", HttpStatusCode.Conflict, Microsoft.ExcelServices.ErrorCodes.InsertDeleteConflict, "stridsBadListInsDel")]
[assembly: HResultError("hrListCannotGrow", HttpStatusCode.Conflict, Microsoft.ExcelServices.ErrorCodes.InsertDeleteConflict, "stridsBadListInsDel")]
[assembly: HResultError("hrNotYetSupportedApiOperation", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.UnsupportedOperation, "stridsApiNotImplemented")]
[assembly: HResultError("SharedInterimIfs::hrRangeSheetsMismatch", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.InvalidArgument, "stridsApiInvalidArgument")]
[assembly: HResultError("SharedInterimIfs::hrRangeParseError", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.InvalidArgument, "stridsApiInvalidArgument")]
[assembly: HResultError("SharedInterimIfs::hrRangeWrong", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.InvalidArgument, "stridsApiInvalidArgument")]
[assembly: HResultError("hrNoPermission", HttpStatusCode.Forbidden, Microsoft.ExcelServices.ErrorCodes.AccessDenied, "stridsApiAccessDenied")]
[assembly: HResultError("E_ACCESSDENIED", HttpStatusCode.Forbidden, Microsoft.ExcelServices.ErrorCodes.AccessDenied, "stridsApiAccessDenied")]
[assembly: HResultError("SharedInterimIfs::hrCreateTableBadListSrcRange", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.InvalidArgument, "stridsBadListPasteSrcRange")]
[assembly: HResultError("SharedInterimIfs::hrGetTableBadListSrcRange", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.InvalidArgument, "stridsBadListSrcRange")]
[assembly: HResultError("SharedInterimIfs::hrCreateTableFormulaInListHdr", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.InvalidArgument, "stridsFormulaInListHdr")]
[assembly: HResultError("SharedInterimIfs::hrCreateTableColHdrTruncate", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.InvalidArgument, "stridsTableColHdrTruncate")]
[assembly: HResultError("SharedInterimIfs::hrGetTableListsOverlap", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.InvalidArgument, "stridsListsOverlap")]
[assembly: HResultError("hrItemAlreadyExists", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.ItemAlreadyExists, "stridsApiItemAlreadyExists")]
[assembly: HResultError("hrNoInterface", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.InvalidArgument, "stridsApiInvalidArgument")]
[assembly: HResultError("DISP_E_UNKNOWNNAME", HttpStatusCode.BadRequest, Microsoft.ExcelServices.ErrorCodes.ApiNotFound, "stridsApiNotFound")]

// Errors 500s
[assembly: HResultError("hrNotImplemented", HttpStatusCode.NotImplemented, Microsoft.ExcelServices.ErrorCodes.NotImplemented, "stridsApiNotImplemented")]
//[assembly: HResultError("hrAborted", HttpStatusCode.InternalServerError, Microsoft.ExcelServices.ErrorCodes.RequestAborted, "stridsApiAborted")] (hrAborted is not yet implemented)

namespace Microsoft.ExcelServices
{
	internal static class ApiSet
	{
		/// <summary>
		/// 1.2 because for now, RTM (1.1) still uses old 16.00 JS file instead of 16.01.
		/// Once redirection is complete, will mark it as 1.1
		/// </summary>
		internal const double PolyfillableDownTo1_1 = 1.2;

		internal static class InProgressFeatures
		{
			internal const double SmallApiAdditions = 1.7;

			// Planned simple APIs, slated likely for 1.5, but still need to be implemented.
			internal const double GetFirstGetLast = 1.8;
			internal const double GetPreviousGetNext = 1.8;
			internal const double WorkbookRange = 1.8;
			internal const double GetSurroundingRegion = 1.8;

			internal const double ChartingApi = 1.9;
		}
	}

	internal static class ErrorCodes
	{
		internal const string GeneralException = "GeneralException";
		internal const string InvalidArgument = "InvalidArgument";
		internal const string InvalidOperation = "InvalidOperation";
		internal const string InvalidSelection = "InvalidSelection";
		internal const string InvalidBinding = "InvalidBinding";
		internal const string InsertDeleteConflict = "InsertDeleteConflict";
		internal const string ItemNotFound = "ItemNotFound";
		internal const string NotImplemented = "NotImplemented";
		internal const string InvalidReference = "InvalidReference";
		internal const string InvalidRequest = "InvalidRequest";
		internal const string ApiNotAvailable = "ApiNotAvailable";
		internal const string Unauthenticated = "Unauthenticated";
		internal const string AccessDenied = "AccessDenied";
		internal const string Conflict = "Conflict";
		internal const string ItemAlreadyExists = "ItemAlreadyExists";
		internal const string ContentLengthRequired = "ContentLengthRequired";
		internal const string ActivityLimitReached = "ActivityLimitReached";
		internal const string RequestAborted = "RequestAborted";
		internal const string ServiceNotAvailable = "ServiceNotAvailable";
		internal const string UnsupportedOperation = "UnsupportedOperation";
		internal const string BadPassword = "BadPassword";
		internal const string ApiNotFound = "ApiNotFound";
	}

	// These need to be defined here since the other file (FunctionsCodeGen.cs) is codegenned
	internal static class FunctionResultDispatchIds
	{
		internal const int FunctionResult_Error = 1;
		internal const int FunctionResult_Value = 2;
	}

#region Event Arguments
	/// <summary>
	/// Provides information about the binding that raised the SelectionChanged event.
	/// </summary>
	[ClientCallableType(ExcludedFromRest = true)]
	[ApiSet(Version = 1.1, IntroducedInVersion = 1.3)]
	public struct BindingSelectionChangedEventArgs
	{
		/// <summary>
		/// Gets the Binding object that represents the binding that raised the SelectionChanged event.
		/// </summary>
		[ApiSet(Version = 1.1, IntroducedInVersion = 1.3)]
		public Binding Binding { get; set; }

		/// <summary>
		/// Gets the index of the first row of the selection (zero-based).
		/// </summary>
		[ApiSet(Version = 1.1, IntroducedInVersion = 1.3)]
		int StartRow { get; set; }
		
		/// <summary>
		/// Gets the index of the first column of the selection (zero-based).
		/// </summary>
		[ApiSet(Version = 1.1, IntroducedInVersion = 1.3)]
		int StartColumn { get; set; }
		
		/// <summary>
		/// Gets the number of rows selected.
		/// </summary>
		[ApiSet(Version = 1.1, IntroducedInVersion = 1.3)]
		int RowCount { get; set; }
		
		/// <summary>
		/// Gets the number of columns selected.
		/// </summary>
		[ApiSet(Version = 1.1, IntroducedInVersion = 1.3)]
		int ColumnCount { get; set; }
	}

	/// <summary>
	/// Provides information about the binding that raised the DataChanged event.
	/// </summary>
	[ClientCallableType(ExcludedFromRest = true)]
	[ApiSet(Version = 1.1, IntroducedInVersion = 1.3)]
	public struct BindingDataChangedEventArgs
	{
		/// <summary>
		/// Gets the Binding object that represents the binding that raised the DataChanged event.
		/// </summary>
		[ApiSet(Version = 1.1, IntroducedInVersion = 1.3)]
		public Binding Binding { get; set; }
	}

	/// <summary>
	/// Provides information about the document that raised the SelectionChanged event.
	/// </summary>
	[ClientCallableType(ExcludedFromRest = true)]
	[ApiSet(Version = 1.1, IntroducedInVersion = 1.3)]
	public struct SelectionChangedEventArgs
	{
		/// <summary>
		/// Gets the workbook object that raised the SelectionChanged event.
		/// </summary>
		[ApiSet(Version = 1.1, IntroducedInVersion = 1.3)]
		public Workbook Workbook { get; set; }
	}

	/// <summary>
	/// Provides information about the setting that raised the SettingsChanged event
	/// </summary>
	[ClientCallableType(ExcludedFromRest = true)]
	[ApiSet(Version = 1.4)]
	public struct SettingsChangedEventArgs
	{
		/// <summary>
		/// Gets the Setting object that represents the binding that raised the SettingsChanged event
		/// </summary>
		[ApiSet(Version = 1.4)]
		public SettingCollection Settings { get; set; }
	}

	#endregion

#region Application
	internal static class ApplicationDispatchIds
	{
		internal const int Application_CalculationMode = 1;
		internal const int Application_Calculate = 2;
		internal const int Application_SuspendApiCalculationUntilNextSync = 3;
	}

	/// <summary>
	/// Represents the Excel application that manages the workbook.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IApplication", InterfaceId = "053AAB3F-C5B6-4A91-93A5-A2C4DA223516", CoClassName = "Application")]
	public interface Application
	{
		/// <summary>
		/// Returns the calculation mode used in the workbook. See Excel.CalculationMode for details. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ApplicationDispatchIds.Application_CalculationMode)]
		CalculationMode CalculationMode { get; }

		/// <summary>
		/// Recalculate all currently opened workbooks in Excel.
		/// </summary>
		/// <param name="calculationType">Specifies the calculation type to use. See Excel.CalculationType for details.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ApplicationDispatchIds.Application_Calculate)]
		void Calculate(CalculationType calculationType);

		/// <summary>
		/// Suspends calculation until the next "context.sync()" is called. Once set, it is the developer's responsibility to re-calc the workbook, to ensure that any dependencies are propagated.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ApplicationDispatchIds.Application_SuspendApiCalculationUntilNextSync)]
		[ClientCallableOperation(RESTfulName = "")]
		void SuspendApiCalculationUntilNextSync();
	}
#endregion Application

#region Workbook
	internal static class WorkbookDispatchIds
	{
		internal const int Workbook_Worksheets = 1;
		internal const int Workbook_Names = 2;
		internal const int Workbook_Tables = 3;
		internal const int Workbook_Application = 4;
		internal const int Workbook_SelectedRange = 5;
		internal const int Workbook_Bindings = 6;
		internal const int Workbook_RemoveReference = 7;
		internal const int Workbook_GetObjectByReferenceId = 8;
		internal const int Workbook_GetObjectTypeNameByReferenceId = 9;
		internal const int Workbook_RemoveAllReferences = 10;
		internal const int Workbook_GetReferenceCount = 11;
		internal const int Workbook_Functions = 12;
		internal const int Workbook_V1Api = 13;
		internal const int Workbook_PivotTables = 14;
		internal const int Workbook_Settings = 15;
		internal const int Workbook_CustomXmlParts = 16;
		internal const int Workbook_ActiveWorksheet = 17;
		internal const int Workbook_GetWorksheetById = 18;
		internal const int Workbook_Range = 19;
		internal const int Workbook_Test = 20;
		internal const int Workbook_Name = 21;
	}

	/// <summary>
	/// Workbook is the top level object which contains related workbook objects such as worksheets, tables, ranges, etc.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IWorkbook", InterfaceId = "bb02266c-6204-4e0d-baa3-cc1a928f573e", CoClassName = "Workbook", ExtensibleObject = true)]
	[ClientCallableServiceRoot]
	public interface Workbook
	{

		/// <summary>
		/// Gets the currently selected range from the workbook.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = WorkbookDispatchIds.Workbook_SelectedRange)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetSelectedRange();

		[ClientCallableComMember(DispatchId = WorkbookDispatchIds.Workbook_GetObjectByReferenceId)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		object _GetObjectByReferenceId(string bstrReferenceId);

		[ClientCallableComMember(DispatchId = WorkbookDispatchIds.Workbook_GetObjectTypeNameByReferenceId)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		string _GetObjectTypeNameByReferenceId(string bstrReferenceId);

		[ClientCallableComMember(DispatchId = WorkbookDispatchIds.Workbook_GetReferenceCount)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		int _GetReferenceCount();

		[ClientCallableComMember(DispatchId = WorkbookDispatchIds.Workbook_RemoveAllReferences)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _RemoveAllReferences();

		[ClientCallableComMember(DispatchId = WorkbookDispatchIds.Workbook_RemoveReference)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _RemoveReference(string bstrReferenceId);

		/// <summary>
		/// Represents Excel application instance that contains this workbook. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = WorkbookDispatchIds.Workbook_Application)]
		Application Application { get; }

		/// <summary>
		/// Represents the collection of custom XML parts contained by this workbook. Read-only.
		/// </summary>
		[ApiSet(Version = 1.5)]
		[ClientCallableComMember(DispatchId = WorkbookDispatchIds.Workbook_CustomXmlParts)]
		CustomXmlPartCollection CustomXmlParts { get; }

		/// <summary>
		/// Represents Excel application instance that contains this workbook. Read-only.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = WorkbookDispatchIds.Workbook_Functions)]
		Functions Functions { get; }

		/// <summary>
		/// Represents a collection of workbook scoped named items (named ranges and constants). Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = WorkbookDispatchIds.Workbook_Names)]
		NamedItemCollection Names { get; }

		/// <summary>
		/// Represents a collection of worksheets associated with the workbook. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = WorkbookDispatchIds.Workbook_Worksheets)]
		WorksheetCollection Worksheets { get; }

		/// <summary>
		/// Represents a collection of tables associated with the workbook. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = WorkbookDispatchIds.Workbook_Tables)]
		TableCollection Tables { get; }

		/// <summary>
		/// Represents a collection of bindings that are part of the workbook. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = WorkbookDispatchIds.Workbook_Bindings)]
		[ClientCallableProperty(ExcludedFromRest = true)]
		BindingCollection Bindings { get; }


		/// <summary>
		/// Represents a collection of PivotTables associated with the workbook. Read-only.
		/// </summary>
		[ApiSet(Version = 1.3)]
		[ClientCallableComMember(DispatchId = WorkbookDispatchIds.Workbook_PivotTables)]
		PivotTableCollection PivotTables { get; }

		/// <summary>
		/// Represents a collection of Settings associated with the workbook. Read-only.
		/// </summary>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = WorkbookDispatchIds.Workbook_Settings)]
		[ClientCallableProperty(ExcludedFromRest = true)]
		SettingCollection Settings { get; }

		/// <summary>
		/// Occurs when the selection in the document is changed.
		/// </summary>
		[ApiSet(Version = 1.1, IntroducedInVersion = 1.3)]
		event EventHandler<SelectionChangedEventArgs> SelectionChanged;

		/// <summary>
		/// For internal use only.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableProperty(ExcludedFromClientLibrary = true, ExcludedFromRest = true)]
		[ClientCallableComMember(DispatchId = WorkbookDispatchIds.Workbook_Test)]
		Test Test { get; }

	}
	#endregion Workbook

	#region Worksheet
	internal static class WorksheetDispatchIds
	{
		internal const int Worksheet_Range = 1;
		internal const int Worksheet_UsedRange = 2;
		internal const int Worksheet_Charts = 3;
		internal const int Worksheet_Cell = 4;
		internal const int Worksheet_Name = 5;
		internal const int Worksheet_Delete = 6;
		internal const int Worksheet_Id = 7;
		internal const int Worksheet_Tables = 8;
		internal const int Worksheet_Activate = 9;
		internal const int Worksheet_Position = 10;
		internal const int Worksheet_OnAccess = 11;
		internal const int Worksheet_Visible = 12;
		internal const int Worksheet_Protection = 13;
		internal const int Worksheet_PivotTables = 14;
		internal const int Worksheet_Names = 15;
		internal const int Worksheet_UsedRangeOrNullObject = 16;
		internal const int Worksheet_RangeByIndexes = 17;
		internal const int Worksheet_Previous = 18;
		internal const int Worksheet_PreviousOrNullObject = 19;
		internal const int Worksheet_Next = 20;
		internal const int Worksheet_NextOrNullObject = 21;
		internal const int Worksheet_TabColor = 22;
		internal const int Worksheet_Calculate = 23;
		internal const int Worksheet_Gridlines = 24;
		internal const int Worksheet_Headings = 25;

		internal const int WorksheetCollection_Indexer = 1;
		internal const int WorksheetCollection_Add = 2;
		internal const int WorksheetCollection_ActiveWorksheet = 3;
		internal const int WorksheetCollection_GetItemOrNullObject = 4;
		internal const int WorksheetCollection_First = 5;
		internal const int WorksheetCollection_Last = 6;
		internal const int WorksheetCollection_GetCount = 7;

		internal const int WorksheetProtection_OnAccess = 1;
		internal const int WorksheetProtection_Protected = 2;
		internal const int WorksheetProtection_Options = 3;
		internal const int WorksheetProtection_Protect = 4;
		internal const int WorksheetProtection_Unprotect = 5;

		internal const int WorksheetProtectionOptions_AllowFormatCells = 1;
		internal const int WorksheetProtectionOptions_AllowFormatColumns = 2;
		internal const int WorksheetProtectionOptions_AllowFormatRows = 3;
		internal const int WorksheetProtectionOptions_AllowInsertColumns = 4;
		internal const int WorksheetProtectionOptions_AllowInsertRows = 5;
		internal const int WorksheetProtectionOptions_AllowInsertHyperlinks = 6;
		internal const int WorksheetProtectionOptions_AllowDeleteColumns = 7;
		internal const int WorksheetProtectionOptions_AllowDeleteRows = 8;
		internal const int WorksheetProtectionOptions_AllowSort = 9;
		internal const int WorksheetProtectionOptions_AllowAutoFilter = 10;
		internal const int WorksheetProtectionOptions_AllowPivotTables = 11;
	}

	/// <summary>
	/// An Excel worksheet is a grid of cells. It can contain data, tables, charts, etc.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableType(DeleteOperationName = "Delete")]
	[ClientCallableComType(Name = "IWorksheet", InterfaceId = "b86e5ae1-476e-4e56-825d-885468e549f3", CoClassName = "Worksheet")]
	public interface Worksheet
	{

		/// <summary>
		/// Activate the worksheet in the Excel UI.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.Worksheet_Activate)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "")]
		void Activate();

		/// <summary>
		/// Calculates all cells on a worksheet.
		/// </summary>
		/// <param name="markAllDirty">Boolean to mark as dirty.</param>	
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.Worksheet_Calculate)]
		void Calculate(bool markAllDirty);

		/// <summary>
		/// Returns collection of charts that are part of the worksheet. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.Worksheet_Charts)]
		ChartCollection Charts { get; }

		/// <summary>
		/// Deletes the worksheet from the workbook.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.Worksheet_Delete)]
		void Delete();

		/// <summary>
		/// Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it's stays within the worksheet grid.
		/// </summary>
		/// <param name="row">The row number of the cell to be retrieved. Zero-indexed.</param>
		/// <param name="column">the column number of the cell to be retrieved. Zero-indexed.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.Worksheet_Cell)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "Cell", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetCell(int row, int column);

		/// <summary>
		/// Gets the range object specified by the address or name.
		/// </summary>
		/// <param name="address">The address or the name of the range. If not specified, the entire worksheet range is returned.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.Worksheet_Range)]
		[ClientCallableOperation(OperationType = OperationType.Read, InvalidateReturnObjectPathAfterRequest = true)]
		Range GetRange([Optional]string address);


		/// <summary>
		/// Gets the worksheet that precedes this one. If there are no previous worksheets, this method will throw an error.
		/// </summary>
		/// <param name="visibleOnly">If true, considers only visible worksheets, skipping over any hidden ones.</param>
		[ApiSet(Version = 1.5)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.Worksheet_Previous)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "Previous", InvalidateReturnObjectPathAfterRequest = true)]
		Worksheet GetPrevious([Optional]bool visibleOnly);

		/// <summary>
		/// Gets the worksheet that precedes this one. If there are no previous worksheets, this method will return a null objet.
		/// </summary>
		/// <param name="visibleOnly">If true, considers only visible worksheets, skipping over any hidden ones.</param>
		[ApiSet(Version = 1.5)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.Worksheet_PreviousOrNullObject)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "", InvalidateReturnObjectPathAfterRequest = true)]
		Worksheet GetPreviousOrNullObject([Optional]bool visibleOnly);

		/// <summary>
		/// Gets the worksheet that follows this one. If there are no worksheets following this one, this method will throw an error.
		/// </summary>
		/// <param name="visibleOnly">If true, considers only visible worksheets, skipping over any hidden ones.</param>
		[ApiSet(Version = 1.5)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.Worksheet_Next)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "Next", InvalidateReturnObjectPathAfterRequest = true)]
		Worksheet GetNext([Optional]bool visibleOnly);

		/// <summary>
		/// Gets the worksheet that follows this one. If there are no worksheets following this one, this method will return a null object.
		/// </summary>
		/// <param name="visibleOnly">If true, considers only visible worksheets, skipping over any hidden ones.</param>
		[ApiSet(Version = 1.5)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.Worksheet_NextOrNullObject)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "", InvalidateReturnObjectPathAfterRequest = true)]
		Worksheet GetNextOrNullObject([Optional]bool visibleOnly);

		/// <summary>
		/// The used range is the smallest range that encompasses any cells that have a value or formatting assigned to them. If the entire worksheet is blank, this function will return the top left cell (i.e.,: it will *not* throw an error).
		/// </summary>
		/// <param name="valuesOnly">Considers only cells with values as used cells (ignoring formatting).</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.Worksheet_UsedRange)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "UsedRange", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetUsedRange([Optional]bool valuesOnly);

		/// <summary>
		/// The used range is the smallest range that encompasses any cells that have a value or formatting assigned to them. If the entire worksheet is blank, this function will return a null object.
		/// </summary>
		/// <param name="valuesOnly">Considers only cells with values as used cells.</param>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.Worksheet_UsedRangeOrNullObject)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetUsedRangeOrNullObject([Optional]bool valuesOnly);

		/// <summary>
		/// Returns a value that uniquely identifies the worksheet in a given workbook. The value of the identifier remains the same even when the worksheet is renamed or moved. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.Worksheet_Id)]
		string Id { get; }

		/// <summary>
		/// The zero-based position of the worksheet within the workbook.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.Worksheet_Position)]
		int Position { get; set; }

		/// <summary>
		/// The display name of the worksheet.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.Worksheet_Name)]
		string Name { get; set; }

		/// <summary>
		/// Collection of tables that are part of the worksheet. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.Worksheet_Tables)]
		TableCollection Tables { get; }

		/// <summary>
		/// The Visibility of the worksheet.
		/// </summary>
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.Worksheet_Visible)]
		[ApiSet(Version = 1.1, CustomText = "1.1 for reading visibility; 1.2 for setting it.")]
		SheetVisibility Visibility { get; set; }

		/// <summary>
		/// Returns sheet protection object for a worksheet.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.Worksheet_Protection)]
		[JsonStringify()]
		WorksheetProtection Protection { get; }

		/// <summary>
		/// Collection of PivotTables that are part of the worksheet. Read-only.
		/// </summary>
		[ApiSet(Version = 1.3)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.Worksheet_PivotTables)]
		PivotTableCollection PivotTables { get; }

		/// <summary>
		/// Collection of names scoped to the current worksheet. Read-only.
		/// </summary>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.Worksheet_Names)]
		NamedItemCollection Names { get; }

	}

	/// <summary>
	/// Represents a collection of worksheet objects that are part of the workbook.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableType(CreateItemOperationName = "Add")]
	[ClientCallableComType(Name = "IWorksheetCollection", InterfaceId = "55a36c77-3310-4afb-aa64-3c1a685f2f50", CoClassName = "WorksheetCollection", SupportEnumeration = true)]
	public interface WorksheetCollection : IEnumerable<Worksheet>
	{
		/// <summary>
		/// Gets the currently active worksheet in the workbook.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.WorksheetCollection_ActiveWorksheet)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "")]
		Worksheet GetActiveWorksheet();

		/// <summary>
		/// Gets a worksheet object using its Name or ID.
		/// </summary>
		/// <param name="key">The Name or ID of the worksheet.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.WorksheetCollection_Indexer)]
		Worksheet this[string key] { get; }

		/// <summary>
		/// Gets a worksheet object using its Name or ID. If the worksheet does not exist, will return a null object.
		/// </summary>
		/// <param name="key">The Name or ID of the worksheet.</param>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.WorksheetCollection_GetItemOrNullObject)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "")]
		Worksheet GetItemOrNullObject(string key);

		/// <summary>
		/// Adds a new worksheet to the workbook. The worksheet will be added at the end of existing worksheets. If you wish to activate the newly added worksheet, call ".activate() on it.
		/// </summary>
		/// <param name="name">The name of the worksheet to be added. If specified, name should be unqiue. If not specified, Excel determines the name of the new worksheet.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.WorksheetCollection_Add)]
		Worksheet Add([Optional]string name);

		/// <summary>
		/// Gets the number of worksheets in the collection.
		/// </summary>
		/// <param name="visibleOnly">Considers only the visible cells.</param>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.WorksheetCollection_GetCount)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		int GetCount([Optional]bool visibleOnly);

		/// <summary>
		/// Gets the first worksheet in the collection.
		/// <param name="visibleOnly">If true, considers only visible worksheets, skipping over any hidden ones.</param>
		/// </summary>
		[ApiSet(Version = 1.5)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.WorksheetCollection_First)]
		[ClientCallableOperation(OperationType = OperationType.Read, InvalidateReturnObjectPathAfterRequest = true)]
		Worksheet GetFirst([Optional]bool visibleOnly);

		/// <summary>
		/// Gets the last worksheet in the collection.
		/// <param name="visibleOnly">If true, considers only visible worksheets, skipping over any hidden ones.</param>
		/// </summary>
		[ApiSet(Version = 1.5)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.WorksheetCollection_Last)]
		[ClientCallableOperation(OperationType = OperationType.Read, InvalidateReturnObjectPathAfterRequest = true)]
		Worksheet GetLast([Optional]bool visibleOnly);
	}

	/// <summary>
	/// Represents the protection of a sheet object.
	/// </summary>
	[ApiSet(Version = 1.2)]
	[ClientCallableComType(Name = "IWorksheetProtection", InterfaceId = "C84C0D35-DEDB-4865-B4A0-B027BAFEC20D", CoClassName = "WorksheetProtection")]
	public interface WorksheetProtection
	{
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.WorksheetProtection_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		/// <summary>
		/// Indicates if the worksheet is protected. Read-Only.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.WorksheetProtection_Protected)]
		bool Protected { get; }
		/// <summary>
		/// Sheet protection options. Read-Only.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.WorksheetProtection_Options)]
		WorksheetProtectionOptions Options { get; }
		/// <summary>
		/// Protects a worksheet. Fails if the worksheet has been protected.
		/// </summary>
		/// <param name="options">sheet protection options.</param>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.WorksheetProtection_Protect)]
		void Protect([Optional]WorksheetProtectionOptions options);
		/// <summary>
		/// Unprotects a worksheet.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.WorksheetProtection_Unprotect)]
		void Unprotect();
	}

	/// <summary>
	/// Represents the options in sheet protection.
	/// </summary>
	[ApiSet(Version = 1.2)]
	[ClientCallableComType(Name = "IWorksheetProtectionOptions", InterfaceId = "201D75BE-81F5-4B2A-A3A8-AE4E72E47ECB", CoClassName = "WorksheetProtectionOptions", CoClassId = "56C94DB3-B781-44CF-9CA8-29FB47A6A267")]
	public struct WorksheetProtectionOptions
	{
		/// <summary>
		/// Represents the worksheet protection option of allowing formatting cells.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.WorksheetProtectionOptions_AllowFormatCells)]
		[Optional]
		bool AllowFormatCells { get; set; }
		/// <summary>
		/// Represents the worksheet protection option of allowing formatting columns.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.WorksheetProtectionOptions_AllowFormatColumns)]
		[Optional]
		bool AllowFormatColumns { get; set; }
		/// <summary>
		/// Represents the worksheet protection option of allowing formatting rows.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.WorksheetProtectionOptions_AllowFormatRows)]
		[Optional]
		bool AllowFormatRows { get; set; }
		/// <summary>
		/// Represents the worksheet protection option of allowing inserting columns.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.WorksheetProtectionOptions_AllowInsertColumns)]
		[Optional]
		bool AllowInsertColumns { get; set; }
		/// <summary>
		/// Represents the worksheet protection option of allowing inserting rows.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.WorksheetProtectionOptions_AllowInsertRows)]
		[Optional]
		bool AllowInsertRows { get; set; }
		/// <summary>
		/// Represents the worksheet protection option of allowing inserting hyperlinks.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.WorksheetProtectionOptions_AllowInsertHyperlinks)]
		[Optional]
		bool AllowInsertHyperlinks { get; set; }
		/// <summary>
		/// Represents the worksheet protection option of allowing deleting columns.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.WorksheetProtectionOptions_AllowDeleteColumns)]
		[Optional]
		bool AllowDeleteColumns { get; set; }
		/// <summary>
		/// Represents the worksheet protection option of allowing deleting rows.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.WorksheetProtectionOptions_AllowDeleteRows)]
		[Optional]
		bool AllowDeleteRows { get; set; }
		/// <summary>
		/// Represents the worksheet protection option of allowing using sort feature.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.WorksheetProtectionOptions_AllowSort)]
		[Optional]
		bool AllowSort { get; set; }
		/// <summary>
		/// Represents the worksheet protection option of allowing using auto filter feature.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.WorksheetProtectionOptions_AllowAutoFilter)]
		[Optional]
		bool AllowAutoFilter { get; set; }
		/// <summary>
		/// Represents the worksheet protection option of allowing using PivotTable feature.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = WorksheetDispatchIds.WorksheetProtectionOptions_AllowPivotTables)]
		[Optional]
		bool AllowPivotTables { get; set; }
	}
#endregion Worksheet

#region Range
	internal static class RangeDispatchIds
	{
		internal const int Range_NumberFormat = 1;  // DO NOT CHANGE Order of NumberFormat and Values
		internal const int Range_Values = 2;        // DO NOT CHANGE Order of NumberFormat and Values
		internal const int Range_Text = 3;
		internal const int Range_Formulas = 4;
		internal const int Range_FormulasLocal = 5;
		internal const int Range_RowIndex = 6;
		internal const int Range_ColumnIndex = 7;
		internal const int Range_RowCount = 8;
		internal const int Range_ColumnCount = 9;
		internal const int Range_Format = 10;
		internal const int Range_Address = 11;
		internal const int Range_AddressLocal = 12;
		internal const int Range_Cell = 13;
		internal const int Range_CellCount = 14;
		internal const int Range_UsedRange = 15;
		internal const int Range_Clear = 16;
		internal const int Range_Insert = 17;
		internal const int Range_Delete = 18;
		internal const int Range_EntireColumn = 19;
		internal const int Range_EntireRow = 20;
		internal const int Range_Worksheet = 21;
		internal const int Range_Select = 22;
		internal const int Range_ReferenceId = 23;
		internal const int Range_KeepReference = 24;
		internal const int Range_GetOffsetRange = 25;
		internal const int Range_GetRow = 26;
		internal const int Range_GetColumn = 27;
		internal const int Range_OnAccess = 28;
		internal const int Range_GetIntersection = 29;
		internal const int Range_GetBoundingRect = 30;
		internal const int Range_ValueTypes = 31;
		internal const int Range_GetLastCell = 32;
		internal const int Range_GetLastColumn = 33;
		internal const int Range_GetLastRow = 34;
		internal const int Range_FormulasR1C1 = 35;
		internal const int Range_Sort = 36;
		internal const int Range_Merge = 37;
		internal const int Range_Unmerge = 38;
		internal const int Range_Hidden = 39;
		internal const int Range_RowHidden = 40;
		internal const int Range_ColumnHidden = 41;
		internal const int Range_ValidateArraySize = 42;
		internal const int Range_GetIntersectionOrNullObject = 43;
		internal const int Range_GetRowsAbove = 44;
		internal const int Range_GetRowsBelow = 45;
		internal const int Range_GetColumnsBefore = 46;
		internal const int Range_GetColumnsAfter = 47;
		internal const int Range_GetResizedRange = 48;
		internal const int Range_RangeView = 49;
		internal const int Range_ConditionalFormats = 50;
		internal const int Range_UsedRangeOrNullObject = 51;
		internal const int Range_SurroundingRegion = 52;
		internal const int Range_isEntireColumn = 53;
		internal const int Range_isEntireRow = 54;
		internal const int Range_Calculate = 55;
		internal const int Range_GetAbsoluteResizedRange = 56;
		internal const int Range_Hyperlink = 57;

		internal const int RangeReference_Address = 1;

		internal const int RangeHyperlink_ScreenTip = 1;
		internal const int RangeHyperlink_Address = 2;
		internal const int RangeHyperlink_DocumentReference = 3;
		internal const int RangeHyperlink_TextToDisplay = 4;
	}

	/// <summary>
	/// Range represents a set of one or more contiguous cells such as a cell, a row, a column, block of cells, etc.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IRange", InterfaceId = "906962e8-a18a-4cc9-9342-279f056bc293", CoClassName = "Range")]
	public interface Range
	{
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		[ApiSet(Version = 1.3)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_ValidateArraySize)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _ValidateArraySize(int rows, int columns);
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_KeepReference)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _KeepReference();
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_ReferenceId)]
		string _ReferenceId { get; }
		/// <summary>
		/// Represents the range reference in A1-style. Address value will contain the Sheet reference (e.g. Sheet1!A1:B4). Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_Address)]
		string Address { get; }
		/// <summary>
		/// Represents range reference for the specified range in the language of the user. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_AddressLocal)]
		string AddressLocal { get; }

		/// <summary>
		/// Calculates a range of cells on a worksheet.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_Calculate)]
		void Calculate();

		/// <summary>
		/// Number of cells in the range. This API will return -1 if the cell count exceeds 2^31-1 (2,147,483,647). Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_CellCount)]
		int CellCount { get; }
		/// <summary>
		/// Clear range values, format, fill, border, etc.
		/// </summary>
		/// <param name="applyTo">Determines the type of clear action. See Excel.ClearApplyTo for details.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_Clear)]
		void Clear([Optional]ClearApplyTo applyTo);
		/// <summary>
		/// Represents the total number of columns in the range. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_ColumnCount)]
		int ColumnCount { get; }
		/// <summary>
		/// Represents the column number of the first cell in the range. Zero-indexed. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_ColumnIndex)]
		int ColumnIndex { get; }
		/// <summary>
		/// Collection of ConditionalFormats that intersect the range. Read-only.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_ConditionalFormats)]
		ConditionalFormatCollection ConditionalFormats { get; }
		/// <summary>
		/// Deletes the cells associated with the range.
		/// </summary>
		/// <param name="shift">Specifies which way to shift the cells. See Excel.DeleteShiftDirection for details.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_Delete)]
		void Delete(DeleteShiftDirection shift);
		/// <summary>
		/// Gets an object that represents the entire column of the range (for example, if the current range represents cells "B4:E11", it's `getEntireColumn` is a range that represents columns "B:E").
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_EntireColumn)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "EntireColumn", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetEntireColumn();
		/// <summary>
		/// Gets an object that represents the entire row of the range (for example, if the current range represents cells "B4:E11", it's `GetEntireRow` is a range that represents rows "4:11").
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_EntireRow)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "EntireRow", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetEntireRow();


		/// <summary>
		/// Represents the type of data of each cell. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_ValueTypes)]
		RangeValueType[][] ValueTypes { get; }
		/// <summary>
		/// Returns a format object, encapsulating the range's font, fill, borders, alignment, and other properties. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_Format)]
		[JsonStringify()]
		RangeFormat Format { get; }
		/// <summary>
		/// Represents the formula in A1-style notation.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_Formulas)]
		object[][] Formulas { get; set; }
		/// <summary>
		/// Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_FormulasLocal)]
		object[][] FormulasLocal { get; set; }
		/// <summary>
		/// Represents the formula in R1C1-style notation.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_FormulasR1C1)]
		object[][] FormulasR1C1 { get; set; }

		/// <summary>
		/// Gets the smallest range object that encompasses the given ranges. For example, the GetBoundingRect of "B2:C5" and "D10:E15" is "B2:E16".
		/// </summary>
		/// <param name="anotherRange">The range object or address or range name.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_GetBoundingRect)]
		[ClientCallableOperation(OperationType = OperationType.Read, InvalidateReturnObjectPathAfterRequest = true)]
		Range GetBoundingRect([TypeScriptType("Excel.Range|string")]object anotherRange);
		/// <summary>
		/// Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it's stays within the worksheet grid. The returned cell is located relative to the top left cell of the range.
		/// </summary>
		/// <param name="row">Row number of the cell to be retrieved. Zero-indexed.</param>
		/// <param name="column">Column number of the cell to be retrieved. Zero-indexed.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_Cell)]
		[ClientCallableOperation(OperationType = OperationType.Read, InvalidateReturnObjectPathAfterRequest = true)]
		Range GetCell(int row, int column);
		/// <summary>
		/// Gets a column contained in the range.
		/// </summary>
		/// <param name="column">Column number of the range to be retrieved. Zero-indexed.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_GetColumn)]
		[ClientCallableOperation(OperationType = OperationType.Read, InvalidateReturnObjectPathAfterRequest = true)]
		Range GetColumn(int column);
		/// <summary>
		/// Gets a certain number of columns to the right of the current Range object.
		/// </summary>
		/// <param name="count">The number of columns to include in the resulting range. In general, use a positive number to create a range outside the current range. You can also use a negative number to create a range within the current range. The default value is 1.</param>
		[ApiSet(Version = 1.1, IntroducedInVersion = 1.3)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "ColumnsAfter", InvalidateReturnObjectPathAfterRequest = true)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_GetColumnsAfter)]
		Range GetColumnsAfter([Optional]int? count);
		/// <summary>
		/// Gets a certain number of columns to the left of the current Range object.
		/// </summary>
		/// <param name="count">The number of columns to include in the resulting range. In general, use a positive number to create a range outside the current range. You can also use a negative number to create a range within the current range. The default value is 1.</param>
		// NOTE: Until implemented in C++, this is an API that is "Polyfill-ed" using JavaScript.  We don't want any codegen for it. Including it here just to capture the signature.
		[ApiSet(Version = 1.1, IntroducedInVersion = 1.3)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "ColumnsBefore", InvalidateReturnObjectPathAfterRequest = true)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_GetColumnsBefore)]
		Range GetColumnsBefore([Optional]int? count);
		/// <summary>
		/// Gets a Range object similar to the current Range object, but with its bottom-right corner expanded (or contracted) by some number of rows and columns.
		/// </summary>
		/// <param name="deltaRows">The number of rows by which to expand the bottom-right corner, relative to the current range. Use a positive number to expand the range, or a negative number to decrease it.</param>
		/// <param name="deltaColumns">The number of columnsby which to expand the bottom-right corner, relative to the current range. Use a positive number to expand the range, or a negative number to decrease it.</param>
		[ApiSet(Version = 1.1, IntroducedInVersion = 1.3)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "ResizedRange", InvalidateReturnObjectPathAfterRequest = true)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_GetResizedRange)]
		Range GetResizedRange(int deltaRows, int deltaColumns);
		/// <summary>
		/// Gets the range object that represents the rectangular intersection of the given ranges.
		/// </summary>
		/// <param name="anotherRange">The range object or range address that will be used to determine the intersection of ranges.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_GetIntersection)]
		[ClientCallableOperation(OperationType = OperationType.Read, InvalidateReturnObjectPathAfterRequest = true)]
		Range GetIntersection([TypeScriptType("Excel.Range|string")]object anotherRange);
		/// <summary>
		/// Gets the range object that represents the rectangular intersection of the given ranges. If no intersection is found, will return a null object.
		/// </summary>
		/// <param name="anotherRange">The range object or range address that will be used to determine the intersection of ranges.</param>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_GetIntersectionOrNullObject)]
		[ClientCallableOperation(OperationType = OperationType.Read, InvalidateReturnObjectPathAfterRequest = true, RESTfulName = "")]
		Range GetIntersectionOrNullObject([TypeScriptType("Excel.Range|string")]object anotherRange);
		/// <summary>
		/// Gets the last cell within the range. For example, the last cell of "B2:D5" is "D5".
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_GetLastCell)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "LastCell", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetLastCell();
		/// <summary>
		/// Gets the last column within the range. For example, the last column of "B2:D5" is "D2:D5".
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_GetLastColumn)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "LastColumn", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetLastColumn();
		/// <summary>
		/// Gets the last row within the range. For example, the last row of "B2:D5" is "B5:D5".
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_GetLastRow)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "LastRow", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetLastRow();
		/// <summary>
		/// Gets an object which represents a range that's offset from the specified range. The dimension of the returned range will match this range. If the resulting range is forced outside the bounds of the worksheet grid, an error will be thrown.
		/// </summary>
		/// <param name="rowOffset">The number of rows (positive, negative, or 0) by which the range is to be offset. Positive values are offset downward, and negative values are offset upward.</param>
		/// <param name="columnOffset">The number of columns (positive, negative, or 0) by which the range is to be offset. Positive values are offset to the right, and negative values are offset to the left.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_GetOffsetRange)]
		[ClientCallableOperation(OperationType = OperationType.Read, InvalidateReturnObjectPathAfterRequest = true)]
		Range GetOffsetRange(int rowOffset, int columnOffset);
		/// <summary>
		/// Gets a row contained in the range.
		/// </summary>
		/// <param name="row">Row number of the range to be retrieved. Zero-indexed.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_GetRow)]
		[ClientCallableOperation(OperationType = OperationType.Read, InvalidateReturnObjectPathAfterRequest = true)]
		Range GetRow(int row);
		/// <summary>
		/// Gets a certain number of rows above the current Range object.
		/// </summary>
		/// <param name="count">The number of rows to include in the resulting range. In general, use a positive number to create a range outside the current range. You can also use a negative number to create a range within the current range. The default value is 1.</param>
		[ApiSet(Version = 1.1, IntroducedInVersion = 1.3)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "RowsAbove", InvalidateReturnObjectPathAfterRequest = true)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_GetRowsAbove)]
		Range GetRowsAbove([Optional]int? count);
		/// <summary>
		/// Gets a certain number of rows below the current Range object.
		/// </summary>
		/// <param name="count">The number of rows to include in the resulting range. In general, use a positive number to create a range outside the current range. You can also use a negative number to create a range within the current range. The default value is 1.</param>
		[ApiSet(Version = 1.1, IntroducedInVersion = 1.3)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "RowsBelow", InvalidateReturnObjectPathAfterRequest = true)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_GetRowsBelow)]
		Range GetRowsBelow([Optional]int? count);
		/// <summary>
		/// Represents if all cells of the current range are hidden.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_Hidden)]
		bool? Hidden { get; }
		/// <summary>
		/// Represents if all rows of the current range are hidden.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_RowHidden)]
		bool? RowHidden { get; set; }
		/// <summary>
		/// Represents if all columns of the current range are hidden.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_ColumnHidden)]
		bool? ColumnHidden { get; set; }
		/// <summary>
		/// Inserts a cell or a range of cells into the worksheet in place of this range, and shifts the other cells to make space. Returns a new Range object at the now blank space.
		/// </summary>
		/// <param name="shift">Specifies which way to shift the cells. See Excel.InsertShiftDirection for details.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_Insert)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true)]
		Range Insert(InsertShiftDirection shift);

		/// <summary>
		/// Merge the range cells into one region in the worksheet.
		/// </summary>
		/// <param name="across">Set true to merge cells in each row of the specified range as separate merged cells. The default value is false.</param> 
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_Merge)]
		void Merge([Optional]bool across);
		/// <summary>
		/// Unmerge the range cells into separate cells.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_Unmerge)]
		void Unmerge();
		/// <summary>
		/// Represents Excel's number format code for the given cell.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_NumberFormat)]
		object[][] NumberFormat { get; set; }
		/// <summary>
		/// Returns the total number of rows in the range. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_RowCount)]
		int RowCount { get; }
		/// <summary>
		/// Returns the row number of the first cell in the range. Zero-indexed. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_RowIndex)]
		int RowIndex { get; }
		/// <summary>
		/// Selects the specified range in the Excel UI.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_Select)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "")]
		void Select();
		/// <summary>
		/// Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_Text)]
		object[][] Text { get; }

		/// <summary>
		/// Returns the used range of the given range object. If there are no used cells within the range, this function will throw an ItemNotFound error.
		/// </summary>
		/// <param name="valuesOnly">Considers only cells with values as used cells.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_UsedRange)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "UsedRange", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetUsedRange([Optional]bool valuesOnly);

		/// <summary>
		/// Returns the used range of the given range object. If there are no used cells within the range, this function will return a null object.
		/// </summary>
		/// <param name="valuesOnly">Considers only cells with values as used cells.</param>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_UsedRangeOrNullObject)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetUsedRangeOrNullObject([Optional]bool valuesOnly);

		/// <summary>
		/// Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_Values)]
		object[][] Values { get; set; }
		/// <summary>
		/// The worksheet containing the current range. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_Worksheet)]
		Worksheet Worksheet { get; }
		/// <summary>
		/// Represents the range sort of the current range.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_Sort)]
		RangeSort Sort { get; }
		/// <summary>
		/// Represents the visible rows of the current range.
		/// </summary>
		[ApiSet(Version = 1.3)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.Range_RangeView)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "VisibleView")]
		RangeView GetVisibleView();



	}

	/// <summary>
	/// Represents a string reference of the form SheetName!A1:B5, or a global or local named range.
	/// </summary>
	[ClientCallableComType(Name = "IRangeReference", InterfaceId = "A253E7A6-82CA-4314-9FEA-411507C37024", CoClassName = "RangeReference", CoClassId = "3A7C6019-23C3-4A18-AEDE-21CD89AAA672")]
	[ApiSet(Version = 1.2)]
	public struct RangeReference
	{
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = RangeDispatchIds.RangeReference_Address)]
		string Address { get; set; }
	}

	#endregion Range

	#region RangeView
	internal static class RangeViewDispatchIds
	{
		internal const int RangeViewCollection_Indexer = 1;
		internal const int RangeViewCollection_First = 2;
		internal const int RangeViewCollection_Last = 3;
		internal const int RangeViewCollection_GetCount = 4;

		internal const int RangeView_OnAccess = 1;
		internal const int RangeView_NumberFormat = 2;    // DO NOT CHANGE Order of NumberFormat and Values
		internal const int RangeView_Values = 3;    // DO NOT CHANGE Order of NumberFormat and Values
		internal const int RangeView_Text = 4;
		internal const int RangeView_Rows = 5;
		internal const int RangeView_Formulas = 6;
		internal const int RangeView_FormulasLocal = 7;
		internal const int RangeView_FormulasR1C1 = 8;
		internal const int RangeView_ValueTypes = 9;
		internal const int RangeView_RowCount = 10;
		internal const int RangeView_ColumnCount = 11;
		internal const int RangeView_Range = 12;
		internal const int RangeView_CellAddresses = 13;
		internal const int RangeView_Index = 14;
		internal const int RangeView_First = 15;
		internal const int RangeView_Last = 16;
	}

	/// <summary>
	/// RangeView represents a set of visible cells of the parent range.
	/// </summary>
	[ApiSet(Version = 1.3)]
	[ClientCallableComType(Name = "IRangeView", InterfaceId = "FE06F84B-2349-433F-B312-A2EFB1BFE2C8", CoClassName = "RangeView")]
	public interface RangeView
	{
		[ClientCallableComMember(DispatchId = RangeViewDispatchIds.RangeView_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Returns a value that represents the index of the RangeView. Read-only.
		/// </summary>
		[ApiSet(Version = 1.3)]
		[ClientCallableComMember(DispatchId = RangeViewDispatchIds.RangeView_Index)]
		int Index { get; }

		/// <summary>
		/// Represents the cell addresses of the RangeView.
		/// </summary>
		[ApiSet(Version = 1.3)]
		[ClientCallableComMember(DispatchId = RangeViewDispatchIds.RangeView_CellAddresses)]
		object[][] CellAddresses { get; }

		/// <summary>
		/// Represents the formula in A1-style notation.
		/// </summary>
		[ApiSet(Version = 1.3)]
		[ClientCallableComMember(DispatchId = RangeViewDispatchIds.RangeView_Formulas)]
		object[][] Formulas { get; set; }

		/// <summary>
		/// Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German.
		/// </summary>
		[ApiSet(Version = 1.3)]
		[ClientCallableComMember(DispatchId = RangeViewDispatchIds.RangeView_FormulasLocal)]
		object[][] FormulasLocal { get; set; }

		/// <summary>
		/// Represents the formula in R1C1-style notation.
		/// </summary>
		[ApiSet(Version = 1.3)]
		[ClientCallableComMember(DispatchId = RangeViewDispatchIds.RangeView_FormulasR1C1)]
		object[][] FormulasR1C1 { get; set; }

		/// <summary>
		/// Represents Excel's number format code for the given cell.
		/// </summary>
		[ApiSet(Version = 1.3)]
		[ClientCallableComMember(DispatchId = RangeViewDispatchIds.RangeView_NumberFormat)]
		object[][] NumberFormat { get; set; }

		/// <summary>
		/// Represents the raw values of the specified range view. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.
		/// </summary>
		[ApiSet(Version = 1.3)]
		[ClientCallableComMember(DispatchId = RangeViewDispatchIds.RangeView_Values)]
		object[][] Values { get; set; }

		/// <summary>
		/// Represents the type of data of each cell. Read-only.
		/// </summary>
		[ApiSet(Version = 1.3)]
		[ClientCallableComMember(DispatchId = RangeViewDispatchIds.RangeView_ValueTypes)]
		RangeValueType[][] ValueTypes { get; }

		/// <summary>
		/// Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API. Read-only.
		/// </summary>
		[ApiSet(Version = 1.3)]
		[ClientCallableComMember(DispatchId = RangeViewDispatchIds.RangeView_Text)]
		object[][] Text { get; }

		/// <summary>
		/// Gets the parent range associated with the current RangeView.
		/// </summary>
		[ApiSet(Version = 1.3)]
		[ClientCallableComMember(DispatchId = RangeViewDispatchIds.RangeView_Range)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "Range", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetRange();

		/// <summary>
		/// Represents a collection of range views associated with the range. Read-only.
		/// </summary>
		[ApiSet(Version = 1.3)]
		[ClientCallableComMember(DispatchId = RangeViewDispatchIds.RangeView_Rows)]
		RangeViewCollection Rows { get; }

		/// <summary>
		/// Returns the number of visible rows. Read-only.
		/// </summary>
		[ApiSet(Version = 1.3)]
		[ClientCallableComMember(DispatchId = RangeViewDispatchIds.RangeView_RowCount)]
		int RowCount { get; }

		/// <summary>
		/// Returns the number of visible columns. Read-only.
		/// </summary>
		[ApiSet(Version = 1.3)]
		[ClientCallableComMember(DispatchId = RangeViewDispatchIds.RangeView_ColumnCount)]
		int ColumnCount { get; }
	}

	/// <summary>
	/// Represents a collection of RangeView objects.
	/// </summary>
	[ApiSet(Version = 1.3)]
	[ClientCallableComType(Name = "IRangeViewCollection", InterfaceId = "BB47319E-6777-4041-B46B-1D6F2AB827A3", CoClassName = "RangeViewCollection", SupportEnumeration = true)]
	public interface RangeViewCollection : IEnumerable<RangeView>
	{
		/// <summary>
		/// Gets a RangeView Row via it's index. Zero-Indexed.
		/// </summary>
		/// <param name="index">Index of the visible row.</param>
		[ApiSet(Version = 1.3)]
		[ClientCallableComMember(DispatchId = RangeViewDispatchIds.RangeViewCollection_Indexer)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		RangeView GetItemAt(int index);

		/// <summary>
		/// Gets the number of RangeView objects in the collection.
		/// </summary>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = RangeViewDispatchIds.RangeViewCollection_GetCount)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		int GetCount();
	}
#endregion

#region Settings
	internal static class SettingsDispatchIds
	{
		internal const int SettingCollection_Indexer = 1;
		internal const int SettingCollection_Set = 2;
		internal const int SettingCollection_Save = 3;
		internal const int SettingCollection_Refresh = 4;
		internal const int SettingCollection_ItemOrNullObject = 5;
		internal const int SettingCollection_GetCount = 6;

		internal const int Setting_OnAccess = 1;
		internal const int Setting_Key = 2;
		internal const int Setting_Value = 3;
		internal const int Setting_Delete = 4;
	}

	/// <summary>
	/// Represents a collection of worksheet objects that are part of the workbook.
	/// </summary>
	[ApiSet(Version = 1.4)]
	[ClientCallableType(ExcludedFromRest = true)]
	[ClientCallableComType(Name = "ISettingCollection", InterfaceId = "4BB24302-09C0-4717-B398-DCC2D834ED4C", CoClassName = "SettingCollection", SupportEnumeration = true)]
	public interface SettingCollection : IEnumerable<Setting>
	{
		/// <summary>
		/// Gets a Setting entry via the key.
		/// </summary>
		/// <param name="key">Key of the setting.</param>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = SettingsDispatchIds.SettingCollection_Indexer)]
		Setting this[string key] { get; }
		/// <summary>
		/// Sets or adds the specified setting to the workbook.
		/// </summary>
		/// <param name="key">The Key of the new setting.</param>
		/// <param name="value">The Value for the new setting.</param>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = SettingsDispatchIds.SettingCollection_Set)]
		Setting Add(string key, [TypeScriptType("string|number|boolean|Array<any>|any")] object value);

		/// <summary>
		/// Gets a Setting entry via the key. If the Setting does not exist, will return a null object.
		/// </summary>
		/// <param name="key">The key of the setting.</param>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = SettingsDispatchIds.SettingCollection_ItemOrNullObject)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "")]
		Setting GetItemOrNullObject(string key);

		/// <summary>
		/// Occurs when the Settings in the document are changed.
		/// </summary>
		[ApiSet(Version = 1.4)]
		event EventHandler<SettingsChangedEventArgs> SettingsChanged;

		/// <summary>
		/// Gets the number of Settings in the collection.
		/// </summary>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = SettingsDispatchIds.SettingCollection_GetCount)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		int GetCount();
	}

	/// <summary>
	/// Setting represents a key-value pair of a setting persisted to the document.
	/// </summary>
	[ApiSet(Version = 1.4)]
	[ClientCallableType(ExcludedFromRest = true)]
	[ClientCallableComType(Name = "ISetting", InterfaceId = "1907D9BB-DED3-498D-BD7C-9EB195333B2C", CoClassName = "Setting")]
	public interface Setting
	{
		[ClientCallableComMember(DispatchId = SettingsDispatchIds.Setting_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Returns the key that represents the id of the Setting. Read-only.
		/// </summary>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = SettingsDispatchIds.Setting_Key)]
		string Key { get; }

		/// <summary>
		/// Represents the value stored for this setting.
		/// </summary>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = SettingsDispatchIds.Setting_Value)]
		object Value { get; set; }

		/// <summary>
		/// Deletes the setting.
		/// </summary>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = SettingsDispatchIds.Setting_Delete)]
		void Delete();
	}
#endregion

#region NamedItem
	internal static class NamedItemDispatchIds
	{
		internal const int NamedItem_Name = 1;
		internal const int NamedItem_Type = 2;
		internal const int NamedItem_Value = 3;
		internal const int NamedItem_Range = 4;
		internal const int NamedItem_Visible = 5;
		internal const int NamedItem_Id = 6;
		internal const int NamedItem_OnAccess = 7;
		internal const int NamedItem_Delete = 8;
		internal const int NamedItem_Comment = 9;
		internal const int NamedItem_RangeOrNull = 10;
		internal const int NamedItem_Scope = 11;
		internal const int NamedItem_Worksheet = 12;
		internal const int NamedItem_WorksheetOrNull = 13;

		internal const int NamedItemCollection_Indexer = 1;
		internal const int NamedItemCollection_GetItemOrNullObject = 2;
		internal const int NamedItemCollection_Add = 3;
		internal const int NamedItemCollection_AddFormulaLocal = 4;
		internal const int NamedItemCollection_OnAccess = 5;
		internal const int NamedItemCollection_GetCount = 6;
	}

	/// <summary>
	/// A collection of all the nameditem objects that are part of the workbook or worksheet, depending on how it was reached.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "INamedItemCollection", InterfaceId = "BD4C9F4B-F762-4779-AF4E-9E9665797830", CoClassName = "NamedItemCollection", SupportEnumeration = true)]
	public interface NamedItemCollection : IEnumerable<NamedItem>
	{
		/// <summary>
		/// Gets a nameditem object using its name
		/// </summary>
		/// <param name="name">nameditem name.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = NamedItemDispatchIds.NamedItemCollection_Indexer)]
		NamedItem this[string name] { get; }

		/// <summary>
		/// Gets a nameditem object using its name. If the nameditem object does not exist, will return a null object.
		/// </summary>
		/// <param name="name">nameditem name.</param>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = NamedItemDispatchIds.NamedItemCollection_GetItemOrNullObject)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "")]
		NamedItem GetItemOrNullObject(string name);
		/// <summary>
		/// Adds a new name to the collection of the given scope.
		/// </summary>
		/// <param name="name">The name of the named item.</param>
		/// <param name="reference">The formula or the range that the name will refer to.</param>
		/// <param name="comment">The comment associated with the named item</param>
		/// <returns></returns>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = NamedItemDispatchIds.NamedItemCollection_Add)]
		NamedItem Add(string name, [TypeScriptType("Excel.Range|string")]object reference, [Optional]string comment);

		/// <summary>
		/// Adds a new name to the collection of the given scope using the user's locale for the formula.
		/// </summary>
		/// <param name="name">The "name" of the named item.</param>
		/// <param name="formula">The formula in the user's locale that the name will refer to.</param>
		/// <param name="comment">The comment associated with the named item</param>
		/// <returns></returns>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = NamedItemDispatchIds.NamedItemCollection_AddFormulaLocal)]
		NamedItem AddFormulaLocal(string name, string formula, [Optional] string comment);

		/// <summary>
		/// Gets the number of named items in the collection.
		/// </summary>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = NamedItemDispatchIds.NamedItemCollection_GetCount)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		int GetCount();

		[ClientCallableComMember(DispatchId = NamedItemDispatchIds.NamedItemCollection_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
	}

	/// <summary>
	/// Represents a defined name for a range of cells or value. Names can be primitive named objects (as seen in the type below), range object, reference to a range. This object can be used to obtain range object associated with names.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableType(DeleteOperationName = "Delete")]
	[ClientCallableComType(Name = "INamedItem", InterfaceId = "E76EE454-3E5E-4187-9389-3C65234609EF", CoClassName = "NamedItem")]
	public interface NamedItem
	{
		[ClientCallableComMember(DispatchId = NamedItemDispatchIds.NamedItem_Id)]
		string _Id { get; }
		[ClientCallableComMember(DispatchId = NamedItemDispatchIds.NamedItem_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		/// <summary>
		/// The name of the object. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = NamedItemDispatchIds.NamedItem_Name)]
		string Name { get; }

		/// <summary>
		/// Returns the range object that is associated with the name. Throws an error if the named item's type is not a range.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = NamedItemDispatchIds.NamedItem_Range)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "Range", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetRange();

		/// <summary>
		/// Returns the range object that is associated with the name. Returns a null object if the named item's type is not a range.
		/// </summary>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = NamedItemDispatchIds.NamedItem_RangeOrNull)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetRangeOrNullObject();

		/// <summary>
		/// Indicates the type of the value returned by the name's formula. See Excel.NamedItemType for details. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = NamedItemDispatchIds.NamedItem_Type)]
		NamedItemType? Type { get; }

		/// <summary>
		/// Represents the value computed by the name's formula. For a named range, will return the range address. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = NamedItemDispatchIds.NamedItem_Value)]
		object Value { get; }

		/// <summary>
		/// Specifies whether the object is visible or not.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = NamedItemDispatchIds.NamedItem_Visible)]
		bool Visible { get; set; }

		/// <summary>
		/// Deletes the given name.
		/// </summary>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = NamedItemDispatchIds.NamedItem_Delete)]
		void Delete();

		/// <summary>
		/// Represents the comment associated with this name.
		/// </summary>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = NamedItemDispatchIds.NamedItem_Comment)]
		string Comment { get; set; }

		/// <summary>
		/// Indicates whether the name is scoped to the workbook or to a specific worksheet. Read-only.
		/// </summary>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = NamedItemDispatchIds.NamedItem_Scope)]
		NamedItemScope Scope { get; }

		/// <summary>
		/// Returns the worksheet on which the named item is scoped to. Throws an error if the items is scoped to the workbook instead.
		/// </summary>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = NamedItemDispatchIds.NamedItem_Worksheet)]
		Worksheet Worksheet { get; }

		/// <summary>
		/// Returns the worksheet on which the named item is scoped to. Returns a null object if the item is scoped to the workbook instead.
		/// </summary>
		[ApiSet(Version = 1.4)]
		[ClientCallableProperty(ExcludedFromRest = true)]
		[ClientCallableComMember(DispatchId = NamedItemDispatchIds.NamedItem_WorksheetOrNull)]
		Worksheet WorksheetOrNullObject { get; }
	}
	#endregion NamedItem

#region Binding
	internal static class BindingDispatchIds
	{
		internal const int Binding_Id = 1;
		internal const int Binding_Type = 2;
		internal const int Binding_Table = 3;
		internal const int Binding_Range = 4;
		internal const int Binding_Text = 5;
		internal const int Binding_OnAccess = 6;
		internal const int Binding_Delete = 7;

		internal const int BindingCollection_Indexer = 1;
		internal const int BindingCollection_Count = 2;
		internal const int BindingCollection_ItemAt = 3;
		internal const int BindingCollection_Add = 4;
		internal const int BindingCollection_AddFromNamedItem = 5;
		internal const int BindingCollection_AddFromSelection = 6;
		internal const int BindingCollection_GetItemOrNullObject = 7;
		internal const int BindingCollection_GetCount = 8;
	}

	/// <summary>
	/// Represents an Office.js binding that is defined in the workbook.
	/// </summary>
	[ClientCallableType(ExcludedFromRest = true)]
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IBinding", InterfaceId = "7957FCE9-D0AF-4302-9F89-6818D8DEC5D5", CoClassName = "Binding")]
	public interface Binding
	{
		[ClientCallableComMember(DispatchId = BindingDispatchIds.Binding_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		/// <summary>
		/// Represents binding identifier. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = BindingDispatchIds.Binding_Id)]
		[ClientCallableProperty(ExcludedFromRest = true)]
		string Id { get; }
		/// <summary>
		/// Deletes the binding.
		/// </summary>
		[ApiSet(Version = 1.3)]
		[ClientCallableComMember(DispatchId = BindingDispatchIds.Binding_Delete)]
		void Delete();
		/// <summary>
		/// Returns the range represented by the binding. Will throw an error if binding is not of the correct type.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = BindingDispatchIds.Binding_Range)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "")]
		Range GetRange();
		/// <summary>
		/// Returns the table represented by the binding. Will throw an error if binding is not of the correct type.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = BindingDispatchIds.Binding_Table)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "")]
		Table GetTable();
		/// <summary>
		/// Returns the text represented by the binding. Will throw an error if binding is not of the correct type.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = BindingDispatchIds.Binding_Text)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "")]
		string GetText();
		/// <summary>
		/// Returns the type of the binding. See Excel.BindingType for details. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = BindingDispatchIds.Binding_Type)]
		[ClientCallableProperty(ExcludedFromRest = true)]
		BindingType Type { get; }

		/// <summary>
		/// Occurs when the selection is changed within the binding.
		/// </summary>
		[ApiSet(Version = 1.1, IntroducedInVersion = 1.3)]
		event EventHandler<BindingSelectionChangedEventArgs> SelectionChanged;

		/// <summary>
		/// Occurs when data or formatting within the binding is changed.
		/// </summary>
		[ApiSet(Version = 1.1, IntroducedInVersion = 1.3)]
		event EventHandler<BindingDataChangedEventArgs> DataChanged;
	}

	/// <summary>
	/// Represents the collection of all the binding objects that are part of the workbook.
	/// </summary>
	[ClientCallableType(ExcludedFromRest = true)]
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IBindingCollection", InterfaceId = "0D1B5A8F-B3C1-4386-A285-5533EA59846E", CoClassName = "BindingCollection")]
	public interface BindingCollection : IEnumerable<Binding>
	{
		/// <summary>
		/// Gets a binding object by ID.
		/// </summary>
		/// <param name="id">Id of the binding object to be retrieved.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = BindingDispatchIds.BindingCollection_Indexer)]
		Binding this[string id] { get; }
		/// <summary>
		/// Gets a binding object by ID. If the binding object does not exist, will return a null object.
		/// </summary>
		/// <param name="id">Id of the binding object to be retrieved.</param>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = BindingDispatchIds.BindingCollection_GetItemOrNullObject)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "")]
		Binding GetItemOrNullObject(string id);
		/// <summary>
		/// Returns the number of bindings in the collection. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = BindingDispatchIds.BindingCollection_Count)]
		int Count { get; }

		/// <summary>
		/// Gets the number of bindings in the collection.
		/// </summary>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = BindingDispatchIds.BindingCollection_GetCount)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		int GetCount();

		/// <summary>
		/// Gets a binding object based on its position in the items array.
		/// </summary>
		/// <param name="index">Index value of the object to be retrieved. Zero-indexed.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = BindingDispatchIds.BindingCollection_ItemAt)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		Binding GetItemAt(int index);

		/// <summary>
		/// Add a new binding to a particular Range.
		/// </summary>
		/// <param name="range">Range to bind the binding to. May be an Excel Range object, or a string. If string, must contain the full address, including the sheet name</param>
		/// <param name="bindingType">Type of binding. See Excel.BindingType.</param>
		/// <param name="id">Name of binding.</param>
		[ApiSet(Version = 1.3)]
		[ClientCallableComMember(DispatchId = BindingDispatchIds.BindingCollection_Add)]
		Binding Add([TypeScriptType("Excel.Range|string")] object range, BindingType bindingType, string id);

		/// <summary>
		/// Add a new binding based on a named item in the workbook.
		/// </summary>
		/// <param name="name">Name from which to create binding.</param>
		/// <param name="bindingType">Type of binding. See Excel.BindingType.</param>
		/// <param name="id">Name of binding.</param>
		[ApiSet(Version = 1.3)]
		[ClientCallableComMember(DispatchId = BindingDispatchIds.BindingCollection_AddFromNamedItem)]
		Binding AddFromNamedItem(string name, BindingType bindingType, string id);

		/// <summary>
		/// Add a new binding based on the current selection.
		/// </summary>
		/// <param name="bindingType">Type of binding. See Excel.BindingType.</param>
		/// <param name="id">Name of binding.</param>
		[ApiSet(Version = 1.3)]
		[ClientCallableComMember(DispatchId = BindingDispatchIds.BindingCollection_AddFromSelection)]
		Binding AddFromSelection(BindingType bindingType, string id);

	}

	#endregion Binding

#region Table
	internal static class TableDispatchIds
	{
		internal const int Table_Id = 1;
		internal const int Table_Name = 2;
		internal const int Table_Range = 3;
		internal const int Table_HeaderRowRange = 4;
		internal const int Table_DataBodyRange = 5;
		internal const int Table_TotalRowRange = 6;
		internal const int Table_ShowHeaders = 7;
		internal const int Table_ShowTotals = 8;
		internal const int Table_TableStyle = 9;
		internal const int Table_TableColumns = 10;
		internal const int Table_TableRows = 11;
		internal const int Table_Delete = 12;
		internal const int Table_OnAccess = 13;
		internal const int Table_Sort = 14;
		internal const int Table_ConvertToRange = 15;
		internal const int Table_Worksheet = 16;
		internal const int Table_ClearFilters = 17;
		internal const int Table_ReapplyFilters = 18;
		internal const int Table_FirstColumn = 19;
		internal const int Table_LastColumn = 20;
		internal const int Table_BandedRows = 21;
		internal const int Table_BandedColumns = 22;
		internal const int Table_FilterButton = 23;

		internal const int TableCollection_Count = 1;
		internal const int TableCollection_Indexer = 2;
		internal const int TableCollection_ItemAt = 3;
		internal const int TableCollection_Add = 4;
		internal const int TableCollection_OnAccess = 5;
		internal const int TableCollection_GetItemOrNullObject = 6;
		internal const int TableCollection_GetCount = 7;

		internal const int TableColumn_Id = 1;
		// = 2 PREVIOUSLY USED ALREADY. DO NOT REUSE THIS ID.
		internal const int TableColumn_Index = 3;
		internal const int TableColumn_Range = 4;
		internal const int TableColumn_HeaderRowRange = 5;
		internal const int TableColumn_DataBodyRange = 6;
		internal const int TableColumn_TotalRowRange = 7;
		internal const int TableColumn_Values = 8;
		internal const int TableColumn_Delete = 9;
		internal const int TableColumn_OnAccess = 10;
		internal const int TableColumn_Filter = 11;
		internal const int TableColumn_Name = 12;
		internal const int TableColumn_Previous = 13;
		internal const int TableColumn_PreviousOrNullObject = 14;
		internal const int TableColumn_Next = 15;
		internal const int TableColumn_NextOrNullObject = 16;

		internal const int TableColumnCollection_Count = 1;
		internal const int TableColumnCollection_Indexer = 2;
		internal const int TableColumnCollection_ItemAt = 3;
		internal const int TableColumnCollection_Insert = 4;
		internal const int TableColumnCollection_OnAccess = 5;
		internal const int TableColumnCollection_GetItemOrNullObject = 6;
		internal const int TableColumnCollection_GetCount = 7;
		internal const int TableColumnCollection_First = 8;
		internal const int TableColumnCollection_Last = 9;

		internal const int TableRow_Index = 1;
		internal const int TableRow_Range = 2;
		internal const int TableRow_Values = 3;
		internal const int TableRow_Delete = 4;
		internal const int TableRow_OnAccess = 5;

		internal const int TableRowCollection_Count = 1;
		internal const int TableRowCollection_ItemAt = 2;
		internal const int TableRowCollection_Insert = 3;
		internal const int TableRowCollection_OnAccess = 4;
		internal const int TableRowCollection_GetCount = 5;
	}

	/// <summary>
	/// Represents a collection of all the tables that are part of the workbook or worksheet, depending on how it was reached.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableType(CreateItemOperationName = "Add")]
	[ClientCallableComType(Name = "ITableCollection", InterfaceId = "D0BDE1B5-7F2E-480A-A803-98CE6BEBB873", CoClassName = "TableCollection")]
	public interface TableCollection : IEnumerable<Table>
	{
		[ClientCallableComMember(DispatchId = TableDispatchIds.TableCollection_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		/// <summary>
		/// Gets a table by Name or ID.
		/// </summary>
		/// <param name="key">Name or ID of the table to be retrieved.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.TableCollection_Indexer)]
		Table this[[TypeScriptType("number|string")]object key] { get; }
		/// <summary>
		/// Gets a table by Name or ID. If the table does not exist, will return a null object.
		/// </summary>
		/// <param name="key">Name or ID of the table to be retrieved.</param>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.TableCollection_GetItemOrNullObject)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "")]
		Table GetItemOrNullObject([TypeScriptType("number|string")]object key);
		
		/// <summary>
		/// Returns the number of tables in the workbook. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.TableCollection_Count)]
		int Count { get; }

		/// <summary>
		/// Gets the number of tables in the collection.
		/// </summary>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.TableCollection_GetCount)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		int GetCount();

		/// <summary>
		/// Gets a table based on its position in the collection.
		/// </summary>
		/// <param name="index">Index value of the object to be retrieved. Zero-indexed.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.TableCollection_ItemAt)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		Table GetItemAt(int index);

		/// <summary>
		/// Create a new table. The range object or source address determines the worksheet under which the table will be added. If the table cannot be added (e.g., because the address is invalid, or the table would overlap with another table), an error will be thrown.
		/// </summary>
		/// <param name="address">A Range object, or a string address or name of the range representing the data source. If the address does not contain a sheet name, the currently-active sheet is used.</param>
		/// <param name="hasHeaders">Boolean value that indicates whether the data being imported has column labels. If the source does not contain headers (i.e,. when this property set to false), Excel will automatically generate header shifting the data down by one row.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.TableCollection_Add)]
		Table Add(object address, bool hasHeaders);
	}

	/// <summary>
	/// Represents an Excel table.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableType(DeleteOperationName = "Delete", ConvertIntegerKeyValueToString = true)]
	[ClientCallableComType(Name = "ITable", InterfaceId = "302DF59F-3294-46A2-8046-6A7647C75847", CoClassName = "Table")]
	public interface Table
	{
		[ClientCallableComMember(DispatchId = TableDispatchIds.Table_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		/// <summary>
		/// Returns a value that uniquely identifies the table in a given workbook. The value of the identifier remains the same even when the table is renamed. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.Table_Id)]
		int Id { get; }
		/// <summary>
		/// Name of the table.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.Table_Name)]
		string Name { get; set; }
		/// <summary>
		/// Gets the range object associated with the entire table.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.Table_Range)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "Range", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetRange();
		/// <summary>
		/// Gets the range object associated with header row of the table.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.Table_HeaderRowRange)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "HeaderRowRange", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetHeaderRowRange();
		/// <summary>
		/// Gets the range object associated with the data body of the table.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.Table_DataBodyRange)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "DataBodyRange", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetDataBodyRange();
		/// <summary>
		/// Gets the range object associated with totals row of the table.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.Table_TotalRowRange)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "TotalRowRange", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetTotalRowRange();
		/// <summary>
		/// Indicates whether the header row is visible or not. This value can be set to show or remove the header row.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.Table_ShowHeaders)]
		bool ShowHeaders { get; set; }
		/// <summary>
		/// Indicates whether the total row is visible or not. This value can be set to show or remove the total row.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.Table_ShowTotals)]
		bool ShowTotals { get; set; }
		/// <summary>
		/// Indicates whether the first column contains special formatting.
		/// </summary>
		[ApiSet(Version = 1.3)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.Table_FirstColumn)]
		bool HighlightFirstColumn { get; set; }
		/// <summary>
		/// Indicates whether the last column contains special formatting.
		/// </summary>
		[ApiSet(Version = 1.3)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.Table_LastColumn)]
		bool HighlightLastColumn { get; set; }
		/// <summary>
		/// Indicates whether the rows show banded formatting in which odd rows are highlighted differently from even ones to make reading the table easier.
		/// </summary>
		[ApiSet(Version = 1.3)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.Table_BandedRows)]
		bool ShowBandedRows { get; set; }
		/// <summary>
		/// Indicates whether the columns show banded formatting in which odd columns are highlighted differently from even ones to make reading the table easier.
		/// </summary>
		[ApiSet(Version = 1.3)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.Table_BandedColumns)]
		bool ShowBandedColumns { get; set; }
		/// <summary>
		/// Indicates whether the filter buttons are visible at the top of each column header. Setting this is only allowed if the table contains a header row.
		/// </summary>
		[ApiSet(Version = 1.3)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.Table_FilterButton)]
		bool ShowFilterButton { get; set; }
		/// <summary>
		/// Constant value that represents the Table style. Possible values are: TableStyleLight1 thru TableStyleLight21, TableStyleMedium1 thru TableStyleMedium28, TableStyleStyleDark1 thru TableStyleStyleDark11. A custom user-defined style present in the workbook can also be specified.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.Table_TableStyle)]
		string Style { get; set; }
		/// <summary>
		/// Represents a collection of all the columns in the table. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.Table_TableColumns)]
		TableColumnCollection Columns { get; }
		/// <summary>
		/// Represents a collection of all the rows in the table. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.Table_TableRows)]
		TableRowCollection Rows { get; }
		/// <summary>
		/// The worksheet containing the current table. Read-only.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.Table_Worksheet)]
		Worksheet Worksheet { get; }
		/// <summary>
		/// Deletes the table.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.Table_Delete)]
		void Delete();
		/// <summary>
		/// Represents the sorting for the table.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.Table_Sort)]
		TableSort Sort { get; }
		/// <summary>
		/// Converts the table into a normal range of cells. All data is preserved.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.Table_ConvertToRange)]
		[ClientCallableOperation(InvalidateReturnObjectPathAfterRequest = true)]
		Range ConvertToRange();

		/// <summary>
		/// Clears all the filters currently applied on the table.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.Table_ClearFilters)]
		void ClearFilters();

		/// <summary>
		/// Reapplies all the filters currently on the table.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.Table_ReapplyFilters)]
		void ReapplyFilters();
	}

	/// <summary>
	/// Represents a collection of all the columns that are part of the table.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableType(CreateItemOperationName = "Add")]
	[ClientCallableComType(Name = "ITableColumnCollection", InterfaceId = "97FD1554-DDA6-49CD-9D39-737AF8297E70", CoClassName = "TableColumnCollection")]
	public interface TableColumnCollection : IEnumerable<TableColumn>
	{
		[ClientCallableComMember(DispatchId = TableDispatchIds.TableColumnCollection_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		/// <summary>
		/// Gets a column object by Name or ID.
		/// </summary>
		/// <param name="key"> Column Name or ID.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.TableColumnCollection_Indexer)]
		TableColumn this[object key] { get; }
		/// <summary>
		/// Gets a column object by Name or ID. If the column does not exist, will return a null object.
		/// </summary>
		/// <param name="key"> Column Name or ID.</param>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.TableColumnCollection_GetItemOrNullObject)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "")]
		TableColumn GetItemOrNullObject(object key);

		/// <summary>
		/// Returns the number of columns in the table. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.TableColumnCollection_Count)]
		int Count { get; }

		/// <summary>
		/// Gets the number of columns in the table.
		/// </summary>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.TableColumnCollection_GetCount)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		int GetCount();

		/// <summary>
		/// Gets a column based on its position in the collection.
		/// </summary>
		/// <param name="index">Index value of the object to be retrieved. Zero-indexed.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.TableColumnCollection_ItemAt)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		TableColumn GetItemAt(int index);

		/// <summary>
		/// Adds a new column to the table.
		/// </summary>
		/// <param name="index">Specifies the relative position of the new column. If null or -1, the addition happens at the end. Columns with a higher index will be shifted to the side. Zero-indexed.</param>
		/// <param name="values">A 2-dimensional array of unformatted values of the table column.</param>
		/// <param name="name">Specifies the name of the new column. If null, the default name will be used.</param>
		[ApiSet(Version = 1.1, CustomText = "1.1 requires an index smaller than the total column count; 1.4 allows index to be optional (null or -1) and will append a column at the end; 1.4 allows name parameter at creation time.")]
		[ClientCallableComMember(DispatchId = TableDispatchIds.TableColumnCollection_Insert)]
		TableColumn Add([Optional]int? index, [Optional]object values, [Optional]string name);
	}

	/// <summary>
	/// Represents a column in a table.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableType(DeleteOperationName = "Delete", ConvertIntegerKeyValueToString = true)]
	[ClientCallableComType(Name = "ITableColumn", InterfaceId = "3291F5CF-437F-482D-BAA1-B0F4C2E430D0", CoClassName = "TableColumn")]
	public interface TableColumn
	{
		[ClientCallableComMember(DispatchId = TableDispatchIds.TableColumn_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		/// <summary>
		/// Returns a unique key that identifies the column within the table. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.TableColumn_Id)]
		int Id { get; }
		/// <summary>
		/// Represents the name of the table column.
		/// </summary>
		[ApiSet(Version = 1.1, CustomText = "1.1 for getting the name; 1.4 for setting it.")]
		[ClientCallableComMember(DispatchId = TableDispatchIds.TableColumn_Name)]
		string Name { get; set; }
		/// <summary>
		/// Returns the index number of the column within the columns collection of the table. Zero-indexed. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.TableColumn_Index)]
		int Index { get; }
		/// <summary>
		/// Gets the range object associated with the entire column.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.TableColumn_Range)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "Range", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetRange();
		/// <summary>
		/// Gets the range object associated with the header row of the column.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.TableColumn_HeaderRowRange)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "HeaderRowRange", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetHeaderRowRange();
		/// <summary>
		/// Gets the range object associated with the data body of the column.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.TableColumn_DataBodyRange)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "DataBodyRange", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetDataBodyRange();
		/// <summary>
		/// Gets the range object associated with the totals row of the column.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.TableColumn_TotalRowRange)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "TotalRowRange", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetTotalRowRange();


		/// <summary>
		/// Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.TableColumn_Values)]
		object[][] Values { get; set; }
		/// <summary>
		/// Deletes the column from the table.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.TableColumn_Delete)]
		void Delete();
		/// <summary>
		/// Retrieve the filter applied to the column.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.TableColumn_Filter)]
		Filter Filter { get; }
	}

	/// <summary>
	/// Represents a collection of all the rows that are part of the table.
	/// 
	/// Note that unlike Ranges or Columns, which will adjust if new rows/columns are added before them,
	/// a TableRow object represent the physical location of the table row, but not the data.
	/// That is, if the data is sorted or if new rows are added, a table row will continue
	/// to point at the index for which it was created.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableType(CreateItemOperationName = "Add")]
	[ClientCallableComType(Name = "ITableRowCollection", InterfaceId = "70544D5B-C1BD-4D4F-A410-87785C4BF2B4", CoClassName = "TableRowCollection")]
	public interface TableRowCollection : IEnumerable<TableRow>
	{
		[ClientCallableComMember(DispatchId = TableDispatchIds.TableRowCollection_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		
		/// <summary>
		/// Returns the number of rows in the table. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.TableRowCollection_Count)]
		int Count { get; }

		/// <summary>
		/// Gets the number of rows in the table.
		/// </summary>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.TableRowCollection_GetCount)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		int GetCount();

		/// <summary>
		/// Gets a row based on its position in the collection.
		/// 
		/// Note that unlike Ranges or Columns, which will adjust if new rows/columns are added before them,
		/// a TableRow object represent the physical location of the table row, but not the data.
		/// That is, if the data is sorted or if new rows are added, a table row will continue
		/// to point at the index for which it was created.
		/// </summary>
		/// <param name="index">Index value of the object to be retrieved. Zero-indexed.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.TableRowCollection_ItemAt)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		TableRow GetItemAt(int index);

		/// <summary>
		/// Adds one or more rows to the table. The return object will be the top of the newly added row(s).
		/// 
		/// Note that unlike Ranges or Columns, which will adjust if new rows/columns are added before them,
		/// a TableRow object represent the physical location of the table row, but not the data.
		/// That is, if the data is sorted or if new rows are added, a table row will continue
		/// to point at the index for which it was created.
		/// </summary>
		/// <param name="index">Specifies the relative position of the new row. If null or -1, the addition happens at the end. Any rows below the inserted row are shifted downwards. Zero-indexed.</param>
		/// <param name="values">A 2-dimensional array of unformatted values of the table row.</param>
		[ApiSet(Version = 1.1, CustomText = "1.1 for adding a single row; 1.4 allows adding of multiple rows.")]
		[ClientCallableComMember(DispatchId = TableDispatchIds.TableRowCollection_Insert)]
		TableRow Add([Optional]int? index, [Optional]object values);
	}

	/// <summary>
	/// Represents a row in a table.
	/// 
	/// Note that unlike Ranges or Columns, which will adjust if new rows/columns are added before them,
	/// a TableRow object represent the physical location of the table row, but not the data.
	/// That is, if the data is sorted or if new rows are added, a table row will continue
	/// to point at the index for which it was created.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableType(DeleteOperationName = "Delete")]
	[ClientCallableComType(Name = "ITableRow", InterfaceId = "2604BD8F-678C-4688-9A24-A43F5B3BE4C2", CoClassName = "TableRow")]
	public interface TableRow
	{
		[ClientCallableComMember(DispatchId = TableDispatchIds.TableRow_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		/// <summary>
		/// Returns the index number of the row within the rows collection of the table. Zero-indexed. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.TableRow_Index)]
		int Index { get; }
		/// <summary>
		/// Returns the range object associated with the entire row.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.TableRow_Range)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "Range", InvalidateReturnObjectPathAfterRequest = true)]
		Range GetRange();
		/// <summary>
		/// Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.TableRow_Values)]
		object[][] Values { get; set; }
		/// <summary>
		/// Deletes the row from the table.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = TableDispatchIds.TableRow_Delete)]
		void Delete();
	}
#endregion Table

#region Range Formats
	internal static class RangeFormatDispatchIds
	{
		internal const int RangeBorder_SideIndex = 1;
		internal const int RangeBorder_LineStyle = 2;
		internal const int RangeBorder_Weight = 3;
		internal const int RangeBorder_Color = 4;
		internal const int RangeBorder_OnAccess = 5;
		internal const int RangeBorder_Id = 6;

		internal const int RangeBorderCollection_Indexer = 1;
		internal const int RangeBorderCollection_Count = 2;
		internal const int RangeBorderCollection_ItemAt = 3;
		internal const int RangeBorderCollection_OnAccess = 4;

		internal const int RangeFill_Color = 1;
		internal const int RangeFill_Clear = 2;
		internal const int RangeFill_OnAccess = 3;

		internal const int RangeFont_Name = 1;
		internal const int RangeFont_Size = 2;
		internal const int RangeFont_Color = 3;
		internal const int RangeFont_Italic = 4;
		internal const int RangeFont_Bold = 5;
		internal const int RangeFont_Underline = 6;
		internal const int RangeFont_OnAccess = 7;

		internal const int RangeFormat_Fill = 1;
		internal const int RangeFormat_Font = 2;
		internal const int RangeFormat_WrapText = 3;
		internal const int RangeFormat_HorizontalAlignment = 4;
		internal const int RangeFormat_VerticalAlignment = 5;
		internal const int RangeFormat_Borders = 6;
		internal const int RangeFormat_OnAccess = 7;
		internal const int RangeFormat_ColumnWidth = 8;
		internal const int RangeFormat_RowHeight = 9;
		internal const int RangeFormat_AutofitColumns = 10;
		internal const int RangeFormat_AutofitRows = 11;
		internal const int RangeFormat_Protection = 12;
		internal const int RangeFormat_TextOrientation = 13;

		internal const int FormatProtection_OnAccess = 1;
		internal const int FormatProtection_Locked = 2;
		internal const int FormatProtection_FormulaHidden = 3;
	}

	/// <summary>
	/// A format object encapsulating the range's font, fill, borders, alignment, and other properties.
	/// </summary>
	[ClientCallableComType(Name = "IRangeFormat", InterfaceId = "E97D0B6E-8FBA-4FD5-9922-495283F3C44C", CoClassName = "RangeFormat")]
	[ApiSet(Version = 1.1)]
	public interface RangeFormat
	{
		[ClientCallableComMember(DispatchId = RangeFormatDispatchIds.RangeFormat_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		/// <summary>
		/// Gets or sets the width of all colums within the range. If the column widths are not uniform, null will be returned.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = RangeFormatDispatchIds.RangeFormat_ColumnWidth)]
		double? ColumnWidth { get; set; }
		/// <summary>
		/// Changes the width of the columns of the current range to achieve the best fit, based on the current data in the columns.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = RangeFormatDispatchIds.RangeFormat_AutofitColumns)]
		void AutofitColumns();
		/// <summary>
		/// Returns the fill object defined on the overall range. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeFormatDispatchIds.RangeFormat_Fill)]
		[JsonStringify()]
		RangeFill Fill { get; }
		/// <summary>
		/// Collection of border objects that apply to the overall range. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeFormatDispatchIds.RangeFormat_Borders)]
		RangeBorderCollection Borders { get; }
		/// <summary>
		/// Returns the font object defined on the overall range. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeFormatDispatchIds.RangeFormat_Font)]
		[JsonStringify()]
		RangeFont Font { get; }
		/// <summary>
		/// Represents the horizontal alignment for the specified object. See Excel.HorizontalAlignment for details.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeFormatDispatchIds.RangeFormat_HorizontalAlignment)]
		HorizontalAlignment? HorizontalAlignment { get; set; }
		/// <summary>
		/// Gets or sets the height of all rows in the range. If the row heights are not uniform null will be returned.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = RangeFormatDispatchIds.RangeFormat_RowHeight)]
		double? RowHeight { get; set; }
		/// <summary>
		/// Changes the height of the rows of the current range to achieve the best fit, based on the current data in the columns.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = RangeFormatDispatchIds.RangeFormat_AutofitRows)]
		void AutofitRows();
		/// <summary>
		/// Represents the vertical alignment for the specified object. See Excel.VerticalAlignment for details.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeFormatDispatchIds.RangeFormat_VerticalAlignment)]
		VerticalAlignment? VerticalAlignment { get; set; }
		/// <summary>
		/// Indicates if Excel wraps the text in the object. A null value indicates that the entire range doesn't have uniform wrap setting
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeFormatDispatchIds.RangeFormat_WrapText)]
		bool? WrapText { get; set; }
		/// <summary>
		/// Returns the format protection object for a range.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = RangeFormatDispatchIds.RangeFormat_Protection)]
		[JsonStringify()]
		FormatProtection Protection { get; }

	}

	/// <summary>
	/// Represents the format protection of a range object.
	/// </summary>
	[ApiSet(Version = 1.2)]
	[ClientCallableComType(Name = "IFormatProtection", InterfaceId = "52AB99FC-FBC1-4E4B-B08B-3AD22314A32E", CoClassName = "FormatProtection")]
	public interface FormatProtection
	{
		[ClientCallableComMember(DispatchId = RangeFormatDispatchIds.FormatProtection_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		/// <summary>
		/// Indicates if Excel locks the cells in the object. A null value indicates that the entire range doesn't have uniform lock setting.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = RangeFormatDispatchIds.FormatProtection_Locked)]
		bool? Locked { get; set; }
		/// <summary>
		/// Indicates if Excel hides the formula for the cells in the range. A null value indicates that the entire range doesn't have uniform formula hidden setting.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = RangeFormatDispatchIds.FormatProtection_FormulaHidden)]
		bool? FormulaHidden { get; set; }
	}

	/// <summary>
	/// Represents the background of a range object.
	/// </summary>
	[ClientCallableComType(Name = "IRangeFill", InterfaceId = "C4514652-D1DB-41D1-8B25-9A27F1B33413", CoClassName = "RangeFill")]
	[ApiSet(Version = 1.1)]
	public interface RangeFill
	{
		[ClientCallableComMember(DispatchId = RangeFormatDispatchIds.RangeFill_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		/// <summary>
		/// HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange")
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeFormatDispatchIds.RangeFill_Color)]
		string Color { get; set; }
		/// <summary>
		/// Resets the range background.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeFormatDispatchIds.RangeFill_Clear)]
		void Clear();
	}

	/// <summary>
	/// Represents the border of an object.
	/// </summary>
	[ClientCallableComType(Name = "IRangeBorder", InterfaceId = "AACFA926-132B-4B49-9D78-1AD4E20B1382", CoClassName = "RangeBorder")]
	[ApiSet(Version = 1.1)]
	public interface RangeBorder
	{
		/// <summary>
		/// Represents border identifier. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = RangeFormatDispatchIds.RangeBorder_Id)]
		[ClientCallableProperty(ExcludedFromClientLibrary = true)]
		[ApiSet(Version = 1.1)]
		BorderIndex Id { get; }

		[ClientCallableComMember(DispatchId = RangeFormatDispatchIds.RangeBorder_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		/// <summary>
		/// HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeFormatDispatchIds.RangeBorder_Color)]
		string Color { get; set; }
		/// <summary>
		/// One of the constants of line style specifying the line style for the border. See Excel.BorderLineStyle for details.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeFormatDispatchIds.RangeBorder_LineStyle)]
		BorderLineStyle? Style { get; set; }
		/// <summary>
		/// Constant value that indicates the specific side of the border. See Excel.BorderIndex for details. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeFormatDispatchIds.RangeBorder_SideIndex)]
		BorderIndex? SideIndex { get; }
		/// <summary>
		/// Specifies the weight of the border around a range. See Excel.BorderWeight for details.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeFormatDispatchIds.RangeBorder_Weight)]
		BorderWeight? Weight { get; set; }
	}

	/// <summary>
	/// Represents the border objects that make up the range border.
	/// </summary>
	[ClientCallableComType(Name = "IRangeBorderCollection", InterfaceId = "BD62C8A4-0125-4EB9-9FE5-91E58E718D06", CoClassName = "RangeBorderCollection")]
	[ApiSet(Version = 1.1)]
	public interface RangeBorderCollection : IEnumerable<RangeBorder>
	{
		[ClientCallableComMember(DispatchId = RangeFormatDispatchIds.RangeBorderCollection_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		/// <summary>
		/// Gets a border object using its name
		/// </summary>
		/// <param name="index">Index value of the border object to be retrieved. See Excel.BorderIndex for details.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeFormatDispatchIds.RangeBorderCollection_Indexer)]
		RangeBorder this[BorderIndex index] { get; }

		/// <summary>
		/// Number of border objects in the collection. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeFormatDispatchIds.RangeBorderCollection_Count)]
		int Count { get; }

		/// <summary>
		/// Gets a border object using its index
		/// </summary>
		/// <param name="index">Index value of the object to be retrieved. Zero-indexed.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeFormatDispatchIds.RangeBorderCollection_ItemAt)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		RangeBorder GetItemAt(int index);
	}

	/// <summary>
	/// This object represents the font attributes (font name, font size, color, etc.) for an object.
	/// </summary>
	[ClientCallableComType(Name = "IRangeFont", InterfaceId = "FAAF874F-30F4-4445-8D6A-F99A6EE81C72", CoClassName = "RangeFont")]
	[ApiSet(Version = 1.1)]
	public interface RangeFont
	{
		[ClientCallableComMember(DispatchId = RangeFormatDispatchIds.RangeFont_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		/// <summary>
		/// Represents the bold status of font.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeFormatDispatchIds.RangeFont_Bold)]
		bool? Bold { get; set; }
		/// <summary>
		/// HTML color code representation of the text color. E.g. #FF0000 represents Red.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeFormatDispatchIds.RangeFont_Color)]
		string Color { get; set; }
		/// <summary>
		/// Represents the italic status of the font.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeFormatDispatchIds.RangeFont_Italic)]
		bool? Italic { get; set; }
		/// <summary>
		/// Font name (e.g. "Calibri")
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeFormatDispatchIds.RangeFont_Name)]
		string Name { get; set; }
		/// <summary>
		/// Font size.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeFormatDispatchIds.RangeFont_Size)]
		double? Size { get; set; }
		/// <summary>
		/// Type of underline applied to the font. See Excel.RangeUnderlineStyle for details.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = RangeFormatDispatchIds.RangeFont_Underline)]
		RangeUnderlineStyle? Underline { get; set; }
	}
#endregion Formats

#region Charts
	internal static class ChartDispatchIds
	{
		internal const int ChartAxes_Category = 1;
		internal const int ChartAxes_Series = 2;
		internal const int ChartAxes_Value = 3;
		internal const int ChartAxes_OnAccess = 4;

		internal const int ChartAxis_MajorGridlines = 1;
		internal const int ChartAxis_MajorUnit = 2;
		internal const int ChartAxis_Maximum = 3;
		internal const int ChartAxis_Minimum = 4;
		internal const int ChartAxis_MinorGridlines = 5;
		internal const int ChartAxis_MinorUnit = 6;
		internal const int ChartAxis_Title = 7;
		internal const int ChartAxis_Format = 8;
		internal const int ChartAxis_Type = 9;
		internal const int ChartAxis_MinorUnitScale = 10;
		internal const int ChartAxis_MajorUnitScale = 11;
		internal const int ChartAxis_BaseUnit = 12;
		internal const int ChartAxis_CategoryNames = 13;
		internal const int ChartAxis_CategoryType = 14;
		internal const int ChartAxis_OnAccess = 15;
		internal const int ChartAxis_DisplayUnit = 16;
		internal const int ChartAxis_ShowDisplayUnitLabel = 17;
		internal const int ChartAxis_CustomDisplayUnit = 18;

		internal const int ChartAxisFormat_Font = 1;
		internal const int ChartAxisFormat_Line = 2;
		internal const int ChartAxisFormat_OnAccess = 3;

		internal const int ChartAxisTitle_Text = 1;
		internal const int ChartAxisTitle_Visible = 2;
		internal const int ChartAxisTitle_Format = 3;
		internal const int ChartAxisTitle_OnAccess = 4;

		internal const int ChartAxisTitleFormat_Font = 1;
		internal const int ChartAxisTitleFormat_OnAccess = 2;

		internal const int Chart_Title = 1;
		internal const int Chart_SetData = 2;
		internal const int Chart_DataLabels = 3;
		internal const int Chart_Legend = 4;
		internal const int Chart_Name = 5;
		internal const int Chart_Top = 6;
		internal const int Chart_Left = 7;
		internal const int Chart_Width = 8;
		internal const int Chart_Height = 9;
		internal const int Chart_Delete = 10;
		internal const int Chart_Series = 11;
		internal const int Chart_Id = 12;
		internal const int Chart_Axes = 13;
		internal const int Chart_Format = 14;
		internal const int Chart_OnAccess = 15;
		internal const int Chart_SetPosition = 16;
		internal const int Chart_GetImage = 17;
		internal const int Chart_Worksheet = 18;

		internal const int ChartAreaFormat_Fill = 1;
		internal const int ChartAreaFormat_Font = 2;
		internal const int ChartAreaFormat_OnAccess = 3;

		internal const int ChartCollection_Add = 1;
		internal const int ChartCollection_Count = 2;
		internal const int ChartCollection_ItemAt = 3;
		internal const int ChartCollection_Indexer = 4;
		internal const int ChartCollection_GetByName = 5;
		internal const int ChartCollection_GetItem = 6;
		internal const int ChartCollection_OnAccess = 7;
		internal const int ChartCollection_GetItemOrNullObject = 8;
		internal const int ChartCollection_GetCount = 9;

		internal const int ChartDataLabels_Position = 1;
		internal const int ChartDataLabels_ShowValue = 2;
		internal const int ChartDataLabels_ShowSeriesName = 3;
		internal const int ChartDataLabels_ShowCategoryName = 4;
		internal const int ChartDataLabels_ShowLegendKey = 5;
		internal const int ChartDataLabels_ShowPercentage = 6;
		internal const int ChartDataLabels_ShowBubbleSize = 7;
		internal const int ChartDataLabels_Separator = 8;
		internal const int ChartDataLabels_Format = 9;
		internal const int ChartDataLabels_OnAccess = 10;

		internal const int ChartDataLabelFormat_Font = 1;
		internal const int ChartDataLabelFormat_Fill = 2;
		internal const int ChartDataLabelFormat_OnAccess = 3;

		internal const int ChartFill_SolidColor = 1;
		internal const int ChartFill_Clear = 2;
		internal const int ChartFill_OnAccess = 3;

		internal const int ChartBorder_SolidColor = 1;
		internal const int ChartBorder_Clear = 2;
		internal const int ChartBorder_OnAccess = 3;

		internal const int ChartFont_Bold = 1;
		internal const int ChartFont_Color = 2;
		internal const int ChartFont_Italic = 3;
		internal const int ChartFont_Name = 4;
		internal const int ChartFont_Size = 5;
		internal const int ChartFont_Underline = 6;
		internal const int ChartFont_OnAccess = 7;

		internal const int ChartGridlines_Visible = 1;
		internal const int ChartGridlines_Format = 2;
		internal const int ChartGridlines_OnAccess = 3;

		internal const int ChartGridlinesFormat_Line = 1;
		internal const int ChartGridlinesFormat_OnAccess = 2;

		internal const int ChartLegend_Visible = 1;
		internal const int ChartLegend_Position = 2;
		internal const int ChartLegend_Overlay = 3;
		internal const int ChartLegend_Format = 4;
		internal const int ChartLegend_OnAccess = 5;

		internal const int ChartLegendFormat_Font = 1;
		internal const int ChartLegendFormat_Fill = 2;
		internal const int ChartLegendFormat_OnAccess = 3;

		internal const int ChartLineFormat_Clear = 1;
		internal const int ChartLineFormat_Color = 2;
		internal const int ChartLineFormat_OnAccess = 3;

		internal const int ChartTitle_Visible = 1;
		internal const int ChartTitle_Text = 2;
		internal const int ChartTitle_Overlay = 3;
		internal const int ChartTitle_Format = 4;
		internal const int ChartTitle_OnAccess = 5;
		internal const int ChartTitle_GetSubstring = 6;
		internal const int ChartTitle_HorizontalAlignment = 7;

		internal const int ChartTitleFormat_Font = 1;
		internal const int ChartTitleFormat_Fill = 2;
		internal const int ChartTitleFormat_OnAccess = 3;

		internal const int ChartPoint_Format = 1;
		internal const int ChartPoint_Value = 2;
		internal const int ChartPoint_OnAccess = 3;

		internal const int ChartPointFormat_Fill = 1;
		internal const int ChartPointFormat_OnAccess = 2;
		internal const int ChartPointFormat_Border = 3;

		internal const int ChartPointsCollection_Count = 1;
		internal const int ChartPointsCollection_ItemAt = 2;
		internal const int ChartPointsCollection_OnAccess = 3;
		internal const int ChartPointsCollection_GetCount = 4;
		internal const int ChartPointsCollection_First = 5;
		internal const int ChartPointsCollection_Last = 6;

		internal const int ChartSeries_Name = 1;
		internal const int ChartSeries_Points = 2;
		internal const int ChartSeries_Format = 3;
		internal const int ChartSeries_OnAccess = 4;
		internal const int ChartSeries_Delete = 5;
		internal const int ChartSeries_SetXAxisValues = 6;
		internal const int ChartSeries_SetValues = 7;
		internal const int ChartSeries_SetBubbleSizes = 8;
		internal const int ChartSeries_Trendlines = 9;

		internal const int ChartSeriesFormat_Fill = 1;
		internal const int ChartSeriesFormat_Line = 2;
		internal const int ChartSeriesFormat_OnAccess = 3;

		internal const int ChartSeriesCollection_Count = 1;
		internal const int ChartSeriesCollection_ItemAt = 2;
		internal const int ChartSeriesCollection_OnAccess = 3;
		internal const int ChartSeriesCollection_GetCount = 4;
		internal const int ChartSeriesCollection_First = 5;
		internal const int ChartSeriesCollection_Last = 6;
		internal const int ChartSeriesCollection_Add = 7;

		internal const int ChartTrendline_Format = 1;
		internal const int ChartTrendline_Type = 2;
		internal const int ChartTrendline_PolynomialOrder = 3;
		internal const int ChartTrendline_MovingAveragePeriod = 4;
		internal const int ChartTrendline_OnAccess = 5;

		internal const int ChartTrendlineCollection_OnAccess = 1;
		internal const int ChartTrendlineCollection_GetCount = 2;
		internal const int ChartTrendlineCollection_Add = 3;

		internal const int ChartTrendlineFormat_Line = 1;
		internal const int ChartTrendlineFormat_OnAccess = 2;

		internal const int ChartFormatString_Font = 1;
		internal const int ChartFormatString_OnAccess = 2;
	}

	/// <summary>
	/// A collection of all the chart objects on a worksheet.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableType(CreateItemOperationName = "Add", HiddenIndexerMethod = true)]
	[ClientCallableComType(Name = "IChartCollection", InterfaceId = "c70eaacf-0ea6-4a54-b148-c600f9a5f5e4", CoClassName = "ChartCollection")]
	public interface ChartCollection : IEnumerable<Chart>
	{
		/// <summary>
		/// Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.
		/// </summary>
		/// <param name="name">Name of the chart to be retrieved.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartCollection_GetItem)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		Chart GetItem(string name);
		/// <summary>
		/// Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.
		/// If the chart does not exist, will return a null object.
		/// </summary>
		/// <param name="name">Name of the chart to be retrieved.</param>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartCollection_GetItemOrNullObject)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "")]
		Chart GetItemOrNullObject(string name);

		/// <summary>
		/// Creates a new chart.
		/// </summary>
		/// <param name="type">Represents the type of a chart. See Excel.ChartType for details.</param>
		/// <param name="sourceData">The Range object corresponding to the source data.</param>
		/// <param name="seriesBy">Specifies the way columns or rows are used as data series on the chart. See Excel.ChartSeriesBy for details.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartCollection_Add)]
		// Note: while sourceData can accept either a Range object or a string (necessary for REST), we will ONLY allow Range objects in JS.
		// Otherwise, desktop code and WAC behavior diverges, given their different treatement of multi-range areas (WAC disallows them), table expansion (desktop does, WAC doesn't), etc.
		Chart Add(ChartType type, object sourceData, [Optional]ChartSeriesBy seriesBy);

		/// <summary>
		/// Returns the number of charts in the worksheet. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartCollection_Count)]
		int Count { get; }

		/// <summary>
		/// Returns the number of charts in the worksheet.
		/// </summary>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartCollection_GetCount)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		int GetCount();

		/// <summary>
		/// Gets a chart based on its position in the collection.
		/// </summary>
		/// <param name="index">Index value of the object to be retrieved. Zero-indexed.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartCollection_ItemAt)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		Chart GetItemAt(int index);
	}

	/// <summary>
	/// Represents a chart object in a workbook.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableType(DeleteOperationName = "Delete")]
	[ClientCallableComType(Name = "IChart", InterfaceId = "b35ce724-5414-4380-8eac-582651db71e7", CoClassName = "Chart")]
	public interface Chart
	{
		[ClientCallableComMember(DispatchId = ChartDispatchIds.Chart_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		[ClientCallableComMember(DispatchId = ChartDispatchIds.Chart_Id)]
		[ClientCallableProperty(ExcludedFromClientLibrary = true)]
		[ApiSet(Version = 1.2)]
		string Id { get; }

		/// <summary>
		/// Represents chart axes. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.Chart_Axes)]
		[JsonStringify()]
		ChartAxes Axes { get; }

		/// <summary>
		/// Represents the datalabels on the chart. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.Chart_DataLabels)]
		[JsonStringify()]
		ChartDataLabels DataLabels { get; }

		/// <summary>
		/// Deletes the chart object.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.Chart_Delete)]
		void Delete();

		/// <summary>
		/// Represents the height, in points, of the chart object.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.Chart_Height)]
		double Height { get; set; }

		/// <summary>
		/// The distance, in points, from the left side of the chart to the worksheet origin.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.Chart_Left)]
		double Left { get; set; }

		/// <summary>
		/// Represents the legend for the chart. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.Chart_Legend)]
		[JsonStringify()]
		ChartLegend Legend { get; }

		/// <summary>
		/// Represents the name of a chart object.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.Chart_Name)]
		string Name { get; set; }

		/// <summary>
		/// Represents either a single series or collection of series in the chart. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.Chart_Series)]
		ChartSeriesCollection Series { get; }

		/// <summary>
		/// Resets the source data for the chart.
		/// </summary>
		/// <param name="sourceData">The Range object corresponding to the source data.</param>
		/// <param name="seriesBy">Specifies the way columns or rows are used as data series on the chart. Can be one of the following: Auto (default), Rows, Columns. See Excel.ChartSeriesBy for details.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.Chart_SetData)]
		// Note: while sourceData can accept either a Range object or a string (necessary for REST), we will ONLY allow Range objects in JS.
		// Otherwise, desktop code and WAC behavior diverges, given their different treatement of multi-range areas (WAC disallows them), table expansion (desktop does, WAC doesn't), etc.
		void SetData(object sourceData, [Optional]ChartSeriesBy seriesBy);

		/// <summary>
		/// Positions the chart relative to cells on the worksheet.
		/// </summary>
		/// <param name="startCell">The start cell. This is where the chart will be moved to. The start cell is the top-left or top-right cell, depending on the user's right-to-left display settings.</param>
		/// <param name="endCell">(Optional) The end cell. If specified, the chart's width and height will be set to fully cover up this cell/range.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.Chart_SetPosition)]
		void SetPosition(object startCell, [Optional]object endCell);

		/// <summary>
		/// Represents the title of the specified chart, including the text, visibility, position and formating of the title. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.Chart_Title)]
		[JsonStringify()]
		ChartTitle Title { get; }

		/// <summary>
		/// Represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.Chart_Top)]
		double Top { get; set; }

		/// <summary>
		/// Represents the width, in points, of the chart object.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.Chart_Width)]
		double Width { get; set; }

		/// <summary>
		/// Encapsulates the format properties for the chart area. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.Chart_Format)]
		[JsonStringify()]
		ChartAreaFormat Format { get; }

		/// <summary>
		/// Renders the chart as a base64-encoded image by scaling the chart to fit the specified dimensions.
		/// The aspect ratio is preserved as part of the resizing.
		/// </summary>
		/// <param name="height">(Optional) The desired height of the resulting image.</param>
		/// <param name="width">(Optional) The desired width of the resulting image.</param>
		/// <param name="fittingMode">(Optional) The method used to scale the chart to the specified to the specified dimensions (if both height and width are set)."</param>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.Chart_GetImage)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "Image")]
		System.IO.Stream GetImage([Optional]int width, [Optional]int height, [Optional]ImageFittingMode fittingMode);

		/// <summary>
		/// The worksheet containing the current chart. Read-only.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.Chart_Worksheet)]
		Worksheet Worksheet { get; }
	}

	/// <summary>
	/// Encapsulates the format properties for the overall chart area.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartAreaFormat", InterfaceId = "8D3ACDD2-720E-4F0D-B318-8EAA58356A9F", CoClassName = "ChartAreaFormat")]
	public interface ChartAreaFormat
	{
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartAreaFormat_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the fill format of an object, which includes background formatting information. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartAreaFormat_Fill)]
		[JsonStringify()]
		ChartFill Fill { get; }

		/// <summary>
		/// Represents the font attributes (font name, font size, color, etc.) for the current object. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartAreaFormat_Font)]
		[JsonStringify()]
		ChartFont Font { get; }
	}

	/// <summary>
	/// Represents a collection of chart series.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartSeriesCollection", InterfaceId = "6FC3E0B3-4A68-4EEE-A181-477EB069BAC1", CoClassName = "ChartSeriesCollection")]
	public interface ChartSeriesCollection : IEnumerable<ChartSeries>
	{
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartSeriesCollection_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Returns the number of series in the collection. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartSeriesCollection_Count)]
		int Count { get; }

		/// <summary>
		/// Add a new series to the collection.
		/// </summary>
		/// <param name="name">Name of the series.</param>
		/// <param name="index">Index value of the series to be added. Zero-indexed.</param>
		[ApiSet(Version = 1.9)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartSeriesCollection_Add)]
		ChartSeries Add([Optional] string name, [Optional] int? index);

		/// <summary>
		/// Returns the number of series in the collection.
		/// </summary>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartSeriesCollection_GetCount)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		int GetCount();

		/// <summary>
		/// Retrieves a series based on its position in the collection.c
		/// </summary>
		/// <param name="index">Index value of the object to be retrieved. Zero-indexed.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartSeriesCollection_ItemAt)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		ChartSeries GetItemAt(int index);


	}

	/// <summary>
	/// Represents a series in a chart.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartSeries", InterfaceId = "54454749-3FDB-401D-B5E6-6667F7F80F11", CoClassName = "ChartSeries")]
	public interface ChartSeries
	{
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartSeries_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the formatting of a chart series, which includes fill and line formatting. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartSeries_Format)]
		[JsonStringify()]
		ChartSeriesFormat Format { get; }

		/// <summary>
		/// Represents the name of a series in a chart.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartSeries_Name)]
		string Name { get; set; }

		/// <summary>
		/// Represents a collection of all points in the series. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartSeries_Points)]
		ChartPointsCollection Points { get; }

		/// <summary>
		/// Represents a collection of Trendlines in the series. Read-only.
		/// </summary>
		[ApiSet(Version = 1.9)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartSeries_Trendlines)]
		ChartTrendlineCollection Trendlines { get; }

		/// <summary>
		/// Deletes the chart series.
		/// </summary>
		[ApiSet(Version = 1.9)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartSeries_Delete)]
		void Delete();

		/// <summary>
		/// Set values of X axis for a chart series. Only works for scatter charts.
		/// </summary>
		/// <param name="sourceData">The Range object corresponding to the source data.</param>
		[ApiSet(Version = 1.9)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartSeries_SetXAxisValues)]
		void SetXAxisValues([TypeScriptType("Excel.Range")]object sourceData);

		/// <summary>
		/// Set values for a chart series. For scatter chart, it means Y axis values.
		/// </summary>
		/// <param name="sourceData">The Range object corresponding to the source data.</param>
		[ApiSet(Version = 1.9)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartSeries_SetValues)]
		void SetValues([TypeScriptType("Excel.Range")]object sourceData);

		/// <summary>
		/// Set bubble sizes for a chart series. Only works for bubble charts.
		/// </summary>
		/// <param name="sourceData">The Range object corresponding to the source data.</param>
		[ApiSet(Version = 1.9)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartSeries_SetBubbleSizes)]
		void SetBubbleSizes([TypeScriptType("Excel.Range")]object sourceData);
	}

	/// <summary>
	/// encapsulates the format properties for the chart series
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartSeriesFormat", InterfaceId = "1D3D150E-E2B2-498C-B53C-57F55E9C6CF6", CoClassName = "ChartSeriesFormat")]
	public interface ChartSeriesFormat
	{
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartSeriesFormat_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the fill format of a chart series, which includes background formating information. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartSeriesFormat_Fill)]
		[JsonStringify()]
		ChartFill Fill { get; }

		/// <summary>
		/// Represents line formatting. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartSeriesFormat_Line)]
		[JsonStringify()]
		ChartLineFormat Line { get; }
	}

	/// <summary>
	/// A collection of all the chart points within a series inside a chart.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartPointsCollection", InterfaceId = "1BDB22BF-3690-4E75-9406-1BF54DB0A127", CoClassName = "ChartPointsCollection")]
	public interface ChartPointsCollection : IEnumerable<ChartPoint>
	{
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartPointsCollection_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Returns the number of chart points in the series. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartPointsCollection_Count)]
		int Count { get; }

		/// <summary>
		/// Returns the number of chart points in the series.
		/// </summary>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartPointsCollection_GetCount)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		int GetCount();

		/// <summary>
		/// Retrieve a point based on its position within the series.
		/// </summary>
		/// <param name="index">Index value of the object to be retrieved. Zero-indexed.</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartPointsCollection_ItemAt)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		ChartPoint GetItemAt(int index);


	}

	/// <summary>
	/// Represents a point of a series in a chart.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartPoint", InterfaceId = "76E71D2A-FB56-4CC8-9375-AFA5C1052E9C", CoClassName = "ChartPoint")]
	public interface ChartPoint
	{
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartPoint_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Encapsulates the format properties chart point. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartPoint_Format)]
		[JsonStringify()]
		ChartPointFormat Format { get; }

		/// <summary>
		/// Returns the value of a chart point. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartPoint_Value)]
		object Value { get; }
	}

	/// <summary>
	/// Represents formatting object for chart points.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartPointFormat", InterfaceId = "D907B031-B51A-4CE6-B903-004554FBD2D2", CoClassName = "ChartPointFormat")]
	public interface ChartPointFormat
	{
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartPointFormat_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the fill format of a chart, which includes background formating information. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartAreaFormat_Fill)]
		[JsonStringify()]
		ChartFill Fill { get; }

		/// <summary>
		/// Represents the border format of a chart point, which includes border formating information. Read-only
		/// </summary>
		[ApiSet(Version = 1.9)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartPointFormat_Border)]
		ChartBorder Border { get; }
	}

	/// <summary>
	/// Represents the chart axes.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartAxes", InterfaceId = "a1635994-4bf2-4358-9a13-924c8ebf53aa", CoClassName = "ChartAxes")]
	public interface ChartAxes
	{
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartAxes_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the category axis in a chart. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartAxes_Category)]
		[JsonStringify()]
		ChartAxis CategoryAxis { get; }

		/// <summary>
		/// Represents the series axis of a 3-dimensional chart. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartAxes_Series)]
		[JsonStringify()]
		ChartAxis SeriesAxis { get; }

		/// <summary>
		/// Represents the value axis in an axis. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartAxes_Value)]
		[JsonStringify()]
		ChartAxis ValueAxis { get; }
	}

	/// <summary>
	/// Represents a single axis in a chart.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartAxis", InterfaceId = "f6beb340-c24b-4087-8127-521e79dc326a", CoClassName = "ChartAxis")]
	public interface ChartAxis
	{
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartAxis_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the formatting of a chart object, which includes line and font formatting. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartAxis_Format)]
		[JsonStringify()]
		ChartAxisFormat Format { get; }

		/// <summary>
		/// Returns a gridlines object that represents the major gridlines for the specified axis. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartAxis_MajorGridlines)]
		[JsonStringify()]
		ChartGridlines MajorGridlines { get; }

		/// <summary>
		/// Represents the interval between two major tick marks. Can be set to a numeric value or an empty string.  The returned value is always a number.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartAxis_MajorUnit)]
		object MajorUnit { get; set; }

		/// <summary>
		/// Represents the maximum value on the value axis.  Can be set to a numeric value or an empty string (for automatic axis values).  The returned value is always a number.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartAxis_Maximum)]
		object Maximum { get; set; }

		/// <summary>
		/// Represents the minimum value on the value axis. Can be set to a numeric value or an empty string (for automatic axis values).  The returned value is always a number.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartAxis_Minimum)]
		object Minimum { get; set; }

		/// <summary>
		/// Returns a Gridlines object that represents the minor gridlines for the specified axis. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartAxis_MinorGridlines)]
		[JsonStringify()]
		ChartGridlines MinorGridlines { get; }

		/// <summary>
		/// Represents the interval between two minor tick marks. "Can be set to a numeric value or an empty string (for automatic axis values). The returned value is always a number.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartAxis_MinorUnit)]
		object MinorUnit { get; set; }

		/// <summary>
		/// Represents the axis title. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartAxis_Title)]
		[JsonStringify()]
		ChartAxisTitle Title { get; }

		/// <summary>
		/// Represents the axis display unit. 
		/// </summary>
		[ApiSet(Version = 1.9)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartAxis_DisplayUnit)]
		ChartAxisDisplayUnit DisplayUnit { get; set; }

		/// <summary>
		/// Represents whether the axis display unit label is visible. 
		/// </summary>
		[ApiSet(Version = 1.9)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartAxis_ShowDisplayUnitLabel)]
		bool ShowDisplayUnitLabel { get; set; }

		/// <summary>
		/// Represents the custom axis display unit value. 
		/// </summary>
		[ApiSet(Version = 1.9)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartAxis_CustomDisplayUnit)]
		double CustomDisplayUnit { get; set; }


		/// <summary>
		/// Represents the axis type. Read-only.
		/// </summary>
		[ApiSet(Version = 1.9)]
		[ClientCallableComMember(DispatchId = ChartDcispatchIds.ChartAxis_Type)]
		AxisType Type { get; }

		/// <summary>
		/// Returns or sets the major unit scale value for the category axis when the CategoryType property is set to TimeScale. 
		/// </summary>
		[ApiSet(Version = 1.9)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartAxis_MajorUnitScale)]
		XlTimeUnit MajorUnitScale { get; set; }

		/// <summary>
		/// Returns or sets the minor unit scale value for the category axis when the CategoryType property is set to TimeScale. 
		/// </summary>
		[ApiSet(Version = 1.9)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartAxis_MinorUnitScale)]
		XlTimeUnit MinorUnitScale { get; set; }

		/// <summary>
		/// Returns or sets the base unit for the specified category axis.
		/// </summary>
		[ApiSet(Version = 1.9)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartAxis_BaseUnit)]
		XlTimeUnit BaseUnit { get; set; }

		/// <summary>
		/// Returns or sets all the category names for the specified axis, as a text array.
		/// </summary>
		[ApiSet(Version = 1.9)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartAxis_CategoryNames)]
		string[] CategoryNames { get; set; }

		/// <summary>
		/// Returns or sets the category axis type.
		/// </summary>
		[ApiSet(Version = 1.9)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartAxis_CategoryType)]
		XlCategoryType CategoryType { get; set; }
	}

	/// <summary>
	/// Encapsulates the format properties for the chart axis.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartAxisFormat", InterfaceId = "3ECEE01A-4340-4F99-82AB-EF9B65646F30", CoClassName = "ChartAxisFormat")]
	public interface ChartAxisFormat
	{
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartAxisFormat_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the font attributes (font name, font size, color, etc.) for a chart axis element. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartAxisFormat_Font)]
		[JsonStringify()]
		ChartFont Font { get; }

		/// <summary>
		/// Represents chart line formatting. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartAxisFormat_Line)]
		[JsonStringify()]
		ChartLineFormat Line { get; }
	}

	/// <summary>
	/// Represents the title of a chart axis.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartAxisTitle", InterfaceId = "ecedd0b6-a619-46f1-bf98-09c97aadd9df", CoClassName = "ChartAxisTitle")]
	public interface ChartAxisTitle
	{
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartAxisTitle_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the formatting of chart axis title. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartAxisTitle_Format)]
		[JsonStringify()]
		ChartAxisTitleFormat Format { get; }

		/// <summary>
		/// Represents the axis title.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartAxisTitle_Text)]
		string Text { get; set; }

		/// <summary>
		/// A boolean that specifies the visibility of an axis title.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartAxisTitle_Visible)]
		bool Visible { get; set; }
	}

	/// <summary>
	/// Represents the chart axis title formatting.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartAxisTitleFormat", InterfaceId = "4CE21BA4-E4C0-4F10-A968-61AFFE7C372F", CoClassName = "ChartAxisTitleFormat")]
	public interface ChartAxisTitleFormat
	{
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartAxisTitleFormat_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the font attributes, such as font name, font size, color, etc. of chart axis title object. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartAxisTitleFormat_Font)]
		[JsonStringify()]
		ChartFont Font { get; }
	}

	/// <summary>
	/// Represents a collection of all the data labels on a chart point.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartDataLabels", InterfaceId = "9fe05b7b-dd28-489d-aab5-7497e4d5c346", CoClassName = "ChartDataLabels")]
	public interface ChartDataLabels
	{
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartDataLabels_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the format of chart data labels, which includes fill and font formatting. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartDataLabels_Format)]
		[JsonStringify()]
		ChartDataLabelFormat Format { get; }

		/// <summary>
		/// DataLabelPosition value that represents the position of the data label. See Excel.ChartDataLabelPosition for details.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartDataLabels_Position)]
		ChartDataLabelPosition? Position { get; set; }

		/// <summary>
		/// Boolean value representing if the data label value is visible or not.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartDataLabels_ShowValue)]
		bool? ShowValue { get; set; }

		/// <summary>
		/// Boolean value representing if the data label series name is visible or not.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartDataLabels_ShowSeriesName)]
		bool? ShowSeriesName { get; set; }

		/// <summary>
		/// Boolean value representing if the data label category name is visible or not.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartDataLabels_ShowCategoryName)]
		bool? ShowCategoryName { get; set; }

		/// <summary>
		/// Boolean value representing if the data label legend key is visible or not.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartDataLabels_ShowLegendKey)]
		bool? ShowLegendKey { get; set; }

		/// <summary>
		/// Boolean value representing if the data label percentage is visible or not.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartDataLabels_ShowPercentage)]
		bool? ShowPercentage { get; set; }

		/// <summary>
		/// Boolean value representing if the data label bubble size is visible or not.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartDataLabels_ShowBubbleSize)]
		bool? ShowBubbleSize { get; set; }

		/// <summary>
		/// String representing the separator used for the data labels on a chart.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartDataLabels_Separator)]
		string Separator { get; set; }
	}

	/// <summary>
	/// Encapsulates the format properties for the chart data labels.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartDataLabelFormat", InterfaceId = "B2BD0519-4F5B-43AC-9584-AD507172CC6F", CoClassName = "ChartDataLabelFormat")]
	public interface ChartDataLabelFormat
	{
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartDataLabelFormat_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the font attributes (font name, font size, color, etc.) for a chart data label. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartDataLabelFormat_Font)]
		[JsonStringify()]
		ChartFont Font { get; }

		/// <summary>
		/// Represents the fill format of the current chart data label. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartDataLabelFormat_Fill)]
		[JsonStringify()]
		ChartFill Fill { get; }
	}

	/// <summary>
	/// Represents major or minor gridlines on a chart axis.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartGridlines", InterfaceId = "7af19b5b-5665-4759-a78e-397318ff75e2", CoClassName = "ChartGridlines")]
	public interface ChartGridlines
	{
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartGridlines_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Boolean value representing if the axis gridlines are visible or not.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartGridlines_Visible)]
		bool Visible { get; set; }

		/// <summary>
		/// Represents the formatting of chart gridlines. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartGridlines_Format)]
		[JsonStringify()]
		ChartGridlinesFormat Format { get; }
	}


	/// <summary>
	/// Encapsulates the format properties for chart gridlines.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartGridlinesFormat", InterfaceId = "DD906913-5D3B-4AC1-88ED-3F2DBC98CB04", CoClassName = "ChartGridlinesFormat")]
	public interface ChartGridlinesFormat
	{
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartGridlinesFormat_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents chart line formatting. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartGridlinesFormat_Line)]
		[JsonStringify()]
		ChartLineFormat Line { get; }
	}

	/// <summary>
	/// Represents the legend in a chart.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartLegend", InterfaceId = "a5c915bf-d752-4b33-95e0-5f84c6e9a46a", CoClassName = "ChartLegend")]
	public interface ChartLegend
	{
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartLegend_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the formatting of a chart legend, which includes fill and font formatting. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartLegend_Format)]
		[JsonStringify()]
		ChartLegendFormat Format { get; }

		/// <summary>
		/// A boolean value the represents the visibility of a ChartLegend object.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartLegend_Visible)]
		bool Visible { get; set; }

		/// <summary>
		/// Represents the position of the legend on the chart. See Excel.ChartLegendPosition for details.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartLegend_Position)]
		ChartLegendPosition? Position { get; set; }

		/// <summary>
		/// Boolean value for whether the chart legend should overlap with the main body of the chart.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartLegend_Overlay)]
		bool? Overlay { get; set; }
	}

	/// <summary>
	/// Encapsulates the format properties of a chart legend.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartLegendFormat", InterfaceId = "B2BD0519-4F5B-43AC-9584-AD507172CC6F", CoClassName = "ChartLegendFormat")]
	public interface ChartLegendFormat
	{
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartLegendFormat_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the font attributes such as font name, font size, color, etc. of a chart legend. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartLegendFormat_Font)]
		[JsonStringify()]
		ChartFont Font { get; }

		/// <summary>
		/// Represents the fill format of an object, which includes background formating information. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartLegendFormat_Fill)]
		[JsonStringify()]
		ChartFill Fill { get; }
	}

	/// <summary>
	/// Represents a chart title object of a chart.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartTitle", InterfaceId = "953ac91f-9c3a-480c-bdec-15d446ad0b82", CoClassName = "ChartTitle")]
	public interface ChartTitle
	{
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartTitle_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the formatting of a chart title, which includes fill and font formatting. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartTitle_Format)]
		[JsonStringify()]
		ChartTitleFormat Format { get; }

		/// <summary>
		/// Boolean value representing if the chart title will overlay the chart or not.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartTitle_Overlay)]
		bool? Overlay { get; set; }

		/// <summary>
		/// Represents the title text of a chart.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartTitle_Text)]
		string Text { get; set; }

		/// <summary>
		/// A boolean value the represents the visibility of a chart title object.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartTitle_Visible)]
		bool Visible { get; set; }

		/// <summary>
		/// Get the characters of a chart title. Line break '\n' also counts one charater.
		/// </summary>
		/// <param name="start">The start index of the sub string</param>
		/// <param name="start">The length of the sub string</param>
		[ApiSet(Version = 1.9)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartTitle_GetSubstring)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		ChartFormatString GetSubstring(int start, int length);

		/// <summary>
		/// Represents the horizontal alignment for chart title.
		/// </summary>
		[ApiSet(Version = 1.9)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartTitle_HorizontalAlignment)]
		ChartTextHorizontalAlignment? HorizontalAlignment { get; set; }
	}

	/// <summary>
	/// Represents the characters in chart object like chart title, chart axis title, etc.
	/// </summary>
	[ApiSet(Version = 1.9)]
	[ClientCallableComType(Name = "IChartFormatString", InterfaceId = "B1AB4E90-1A7D-4BEF-897E-FEB990ABC4B7", CoClassName = "ChartFormatString")]
	public interface ChartFormatString
	{
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartFormatString_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the font attributes, such as font name, font size, color, etc. of chart characters object. Read-only.
		/// </summary>
		[ApiSet(Version = 1.9)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartFormatString_Font)]
		ChartFont Font { get; }
	}

	/// <summary>
	/// Provides access to the office art formatting for chart title.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartTitleFormat", InterfaceId = "ACA6BCFA-EFDD-4B81-9478-B1508EC42CB9", CoClassName = "ChartTitleFormat")]
	public interface ChartTitleFormat
	{
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartTitleFormat_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the font attributes (font name, font size, color, etc.) for an object. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartTitleFormat_Font)]
		[JsonStringify()]
		ChartFont Font { get; }

		/// <summary>
		/// Represents the fill format of an object, which includes background formating information. Read-only.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartTitleFormat_Fill)]
		[JsonStringify()]
		ChartFill Fill { get; }
	}

	/// <summary>
	/// Represents the fill formatting for a chart element.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartFill", InterfaceId = "3147230c-d46d-40ea-b3d8-11970eb8a0af", CoClassName = "ChartFill")]
	public interface ChartFill
	{
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartFill_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Clear the fill color of a chart element.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartFill_Clear)]
		void Clear();

		/// <summary>
		/// Sets the fill formatting of a chart element to a uniform color.
		/// </summary>
		/// <param name="color">HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").</param>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartFill_SolidColor)]
		void SetSolidColor(string color);
	}

	/// <summary>
	/// Represents the border formatting for a chart element.
	/// </summary>
	[ApiSet(Version = 1.9)]
	[ClientCallableComType(Name = "IChartBorder", InterfaceId = "acdf7ab0-69e6-4c93-9a52-74e16a42b36a", CoClassName="ChartBorder")]
	public interface ChartBorder
	{
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartBorder_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Clear the border color of a chart element.
		/// </summary>
		[ApiSet(Version = 1.9)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartBorder_Clear)]
		void Clear();

		/// <summary>
		/// Sets the border formatting of a chart element to a uniform color.
		/// </summary>
		/// <param name="color">HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").</param>
		[ApiSet(Version = 1.9)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartBorder_SolidColor)]
		void SetSolidColor(string color);
	}

	/// <summary>
	/// Enapsulates the formatting options for line elements.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartLineFormat", InterfaceId = "0E0D5F3D-DB8D-46CC-B268-BA1D3D190A38", CoClassName = "ChartLineFormat")]
	public interface ChartLineFormat
	{
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartLineFormat_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Clear the line format of a chart element.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartLineFormat_Clear)]
		void Clear();

		/// <summary>
		/// HTML color code representing the color of lines in the chart.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartLineFormat_Color)]
		string Color { get; set; }
	}

	/// <summary>
	/// This object represents the font attributes (font name, font size, color, etc.) for a chart object.
	/// </summary>
	[ApiSet(Version = 1.1)]
	[ClientCallableComType(Name = "IChartFont", InterfaceId = "d62d7af0-54f2-4c16-9e0b-8d5a0ff611b2", CoClassName = "ChartFont")]
	public interface ChartFont
	{
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartFont_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the bold status of font.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartFont_Bold)]
		bool? Bold { get; set; }

		/// <summary>
		/// HTML color code representation of the text color. E.g. #FF0000 represents Red.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartFont_Color)]
		string Color { get; set; }

		/// <summary>
		/// Represents the italic status of the font.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartFont_Italic)]
		bool? Italic { get; set; }

		/// <summary>
		/// Font name (e.g. "Calibri")
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartFont_Name)]
		string Name { get; set; }

		/// <summary>
		/// Size of the font (e.g. 11)
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartFont_Size)]
		double? Size { get; set; }

		/// <summary>
		/// Type of underline applied to the font. See Excel.ChartUnderlineStyle for details.
		/// </summary>
		[ApiSet(Version = 1.1)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartFont_Underline)]
		ChartUnderlineStyle? Underline { get; set; }
	}

	/// <summary>
	/// This object represents the attributes for a chart trendline object.
	/// </summary>
	[ApiSet(Version = 1.9)]
	[ClientCallableComType(Name = "IChartTrendline", InterfaceId = "B0AB4E90-1A7D-4BEF-897E-FEB990ABC4B7", CoClassName = "ChartTrendline")]
	public interface ChartTrendline
	{
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartTrendline_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the formatting of a chart trendline. Read-only.
		/// </summary>
		[ApiSet(Version = 1.9)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartTrendline_Format)]
		[JsonStringify()]
		ChartTrendlineFormat Format { get; }

		/// <summary>
		/// Represents the Type of a chart trendline.
		/// </summary>
		[ApiSet(Version = 1.9)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartTrendline_Type)]
		TrendlineType Type { get; set; }

		/// <summary>
		/// Represents the PolynomialOrder of a chart trendline, specific for trendline with Polynomial type.
		/// </summary>
		[ApiSet(Version = 1.9)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartTrendline_PolynomialOrder)]
		int? PolynomialOrder { get; set; }

		/// <summary>
		/// Represents the MovingAveragePeriod of a chart trendline, specific for trendline with MovingAverage type.
		/// </summary>
		[ApiSet(Version = 1.9)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartTrendline_MovingAveragePeriod)]
		int? MovingAveragePeriod { get; set; }
	}

	/// <summary>
	/// Represents a collection of Chart Trendlines.
	/// </summary>
	[ApiSet(Version = 1.9)]
	[ClientCallableComType(Name = "IChartTrendlineCollection", InterfaceId = "E8291097-2B13-414B-AC8D-5C2CD460BCEF", CoClassName = "ChartTrendlineCollection")]
	public interface ChartTrendlineCollection : IEnumerable<ChartTrendline>
	{
		/// <summary>
		/// Returns the number of trendlines in the collection.
		/// </summary>
		[ApiSet(Version = 1.9)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartTrendlineCollection_GetCount)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		int GetCount();

		/// <summary>
		/// Adds a new trendline to trendline collection.
		/// </summary>
		/// <param name="type">TrendlineType.</param>
		[ApiSet(Version = 1.9)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartTrendlineCollection_Add)]
		ChartTrendline Add([Optional] TrendlineType type);

		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartTrendlineCollection_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
	}


	/// <summary>
	/// Encapsulates the format properties for chart trendline.
	/// </summary>
	[ApiSet(Version = 1.9)]
	[ClientCallableComType(Name = "IChartTrendlineFormat", InterfaceId = "DD906913-5D3B-4AC1-88ED-3F2DBC98CB03", CoClassName = "ChartTrendlineFormat")]
	public interface ChartTrendlineFormat
	{
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartTrendlineFormat_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents chart line formatting. Read-only.
		/// </summary>
		[ApiSet(Version = 1.9)]
		[ClientCallableComMember(DispatchId = ChartDispatchIds.ChartTrendlineFormat_Line)]
		[JsonStringify()]
		ChartLineFormat Line { get; }
	}

	#endregion Charts

#region Sort
	internal static class SortDispatchIds
	{
		internal const int RangeSort_Apply = 1;
		internal const int RangeSort_OnAccess = 2;

		internal const int SortField_Key = 1;
		internal const int SortField_SortOn = 2;
		internal const int SortField_Ascending = 3;
		internal const int SortField_Color = 4;
		internal const int SortField_DataOption = 5;
		internal const int SortField_Icon = 6;

		internal const int TableSort_Apply = 1;
		internal const int TableSort_MatchCase = 2;
		internal const int TableSort_Method = 3;
		internal const int TableSort_OnAccess = 4;
		internal const int TableSort_Clear = 5;
		internal const int TableSort_Reapply = 6;
		internal const int TableSort_Fields = 7;
	}

	/// <summary>
	/// Manages sorting operations on Range objects.
	/// </summary>
	[ApiSet(Version = 1.2)]
	[ClientCallableComType(Name = "IRangeSort", InterfaceId = "8D69987D-B7AD-4AF7-B297-529C21A39ACC", CoClassName = "RangeSort")]
	public interface RangeSort
	{

		[ClientCallableComMember(DispatchId = SortDispatchIds.RangeSort_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Perform a sort operation.
		/// </summary>
		/// <param name="fields">The list of conditions to sort on.</param>
		/// <param name="matchCase">Whether to have the casing impact string ordering.</param>
		/// <param name="hasHeaders">Whether the range has a header.</param>
		/// <param name="orientation">Whether the operation is sorting rows or columns.</param>
		/// <param name="method">The ordering method used for Chinese characters.</param>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = SortDispatchIds.RangeSort_Apply)]
		void Apply(SortField[] fields, [Optional] bool matchCase, [Optional] bool hasHeaders, [Optional] SortOrientation orientation, [Optional] SortMethod method);
	}

	/// <summary>
	/// Manages sorting operations on Table objects.
	/// </summary>
	[ApiSet(Version = 1.2)]
	[ClientCallableComType(Name = "ITableSort", InterfaceId = "2FA61C80-F2B7-46A2-8713-AD13E8C3DC4E", CoClassName = "TableSort")]
	public interface TableSort
	{
		[ClientCallableComMember(DispatchId = SortDispatchIds.TableSort_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Perform a sort operation.
		/// </summary>
		/// <param name="fields">The list of conditions to sort on.</param>
		/// <param name="matchCase">Whether to have the casing impact string ordering.</param>
		/// <param name="method">The ordering method used for Chinese characters.</param>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = SortDispatchIds.TableSort_Apply)]
		void Apply(SortField[] fields, [Optional] bool matchCase, [Optional] SortMethod method);

		/// <summary>
		/// Represents whether the casing impacted the last sort of the table.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = SortDispatchIds.TableSort_MatchCase)]
		bool MatchCase { get; }

		/// <summary>
		/// Represents Chinese character ordering method last used to sort the table.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = SortDispatchIds.TableSort_Method)]
		SortMethod Method { get; }

		/// <summary>
		/// Clears the sorting that is currently on the table. While this doesn't modify the table's ordering, it clears the state of the header buttons.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = SortDispatchIds.TableSort_Clear)]
		void Clear();

		/// <summary>
		/// Reapplies the current sorting parameters to the table.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = SortDispatchIds.TableSort_Reapply)]
		void Reapply();

		/// <summary>
		/// Represents the current conditions used to last sort the table.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = SortDispatchIds.TableSort_Fields)]
		SortField[] Fields { get; }
	}

	/// <summary>
	/// Represents a condition in a sorting operation.
	/// </summary>
	[ApiSet(Version = 1.2)]
	[ClientCallableComType(Name = "ISortField", InterfaceId = "DFE7801F-F972-476C-A4ED-1E6E11D59148", CoClassName = "SortField", CoClassId = "9EB4FF82-6464-49F8-908A-A744F934AB17")]
	public struct SortField
	{
		/// <summary>
		/// Represents the column (or row, depending on the sort orientation) that the condition is on. Represented as an offset from the first column (or row).
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = SortDispatchIds.SortField_Key)]
		int Key { get; set; }

		/// <summary>
		/// Represents the type of sorting of this condition.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = SortDispatchIds.SortField_SortOn)]
		[Optional]
		SortOn SortOn { get; set; }

		/// <summary>
		/// Represents whether the sorting is done in an ascending fashion.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = SortDispatchIds.SortField_Ascending)]
		[Optional]
		bool Ascending { get; set; }

		/// <summary>
		/// Represents the color that is the target of the condition if the sorting is on font or cell color.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = SortDispatchIds.SortField_Color)]
		[Optional]
		string Color { get; set; }

		/// <summary>
		/// Represents additional sorting options for this field.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = SortDispatchIds.SortField_DataOption)]
		[Optional]
		SortDataOption DataOption { get; set; }

		/// <summary>
		/// Represents the icon that is the target of the condition if the sorting is on the cell's icon.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = SortDispatchIds.SortField_Icon)]
		[Optional]
		Icon Icon { get; set; }
	}
#endregion Sort

#region Filter
	internal static class FilterDispatchIds
	{
		internal const int Filter_Apply = 1;
		internal const int Filter_OnAccess = 2;
		internal const int Filter_Clear = 3;
		internal const int Filter_Criteria = 4;
		internal const int Filter_BottomItems = 5;
		internal const int Filter_BottomPercent = 6;
		internal const int Filter_CellColor = 7;
		internal const int Filter_Dynamic = 8;
		internal const int Filter_FontColor = 9;
		internal const int Filter_Values = 10;
		internal const int Filter_TopItems = 11;
		internal const int Filter_TopPercent = 12;
		internal const int Filter_Icon = 13;
		internal const int Filter_Custom = 14;

		internal const int FilterCriteria_Criterion1 = 1;
		internal const int FilterCriteria_Criterion2 = 2;
		internal const int FilterCriteria_Color = 3;
		internal const int FilterCriteria_Operator = 4;
		internal const int FilterCriteria_Icon = 5;
		internal const int FilterCriteria_DynamicCriteria = 6;
		internal const int FilterCriteria_Values = 7;
		internal const int FilterCriteria_FilterOn = 8;

		internal const int FilterDatetime_Date = 1;
		internal const int FilterDatetime_Specificity = 2;
	}

	/// <summary>
	/// Manages the filtering of a table's column.
	/// </summary>
	[ApiSet(Version = 1.2)]
	[ClientCallableComType(Name = "IFilter", InterfaceId = "44E193B3-7AE0-4F97-9A63-D79033780ECF", CoClassName = "Filter")]
	public interface Filter
	{
		[ClientCallableComMember(DispatchId = FilterDispatchIds.Filter_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Apply the given filter criteria on the given column.
		/// </summary>
		/// <param name="criteria">The criteria to apply.</param>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = FilterDispatchIds.Filter_Apply)]
		void Apply(FilterCriteria criteria);

		/// <summary>
		/// Apply a "Bottom Item" filter to the column for the given number of elements.
		/// </summary>
		/// <param name="count">The number of elements from the bottom to show.</param>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = FilterDispatchIds.Filter_BottomItems)]
		void ApplyBottomItemsFilter(int count);

		/// <summary>
		/// Apply a "Bottom Percent" filter to the column for the given percentage of elements.
		/// </summary>
		/// <param name="percent">The percentage of elements from the bottom to show.</param>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = FilterDispatchIds.Filter_BottomPercent)]
		void ApplyBottomPercentFilter(int percent);

		/// <summary>
		/// Apply a "Cell Color" filter to the column for the given color.
		/// </summary>
		/// <param name="color">The background color of the cells to show.</param>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = FilterDispatchIds.Filter_CellColor)]
		void ApplyCellColorFilter(string color);

		/// <summary>
		/// Apply a "Dynamic" filter to the column.
		/// </summary>
		/// <param name="criteria">The dynamic criteria to apply.</param>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = FilterDispatchIds.Filter_Dynamic)]
		void ApplyDynamicFilter(DynamicFilterCriteria criteria);

		/// <summary>
		/// Apply a "Font Color" filter to the column for the given color.
		/// </summary>
		/// <param name="color">The font color of the cells to show.</param>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = FilterDispatchIds.Filter_FontColor)]
		void ApplyFontColorFilter(string color);

		/// <summary>
		/// Apply a "Values" filter to the column for the given values.
		/// </summary>
		/// <param name="values">The list of values to show.</param>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = FilterDispatchIds.Filter_Values)]
		void ApplyValuesFilter(object[] values);

		/// <summary>
		/// Apply a "Top Item" filter to the column for the given number of elements.
		/// </summary>
		/// <param name="count">The number of elements from the top to show.</param>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = FilterDispatchIds.Filter_TopItems)]
		void ApplyTopItemsFilter(int count);

		/// <summary>
		/// Apply a "Top Percent" filter to the column for the given percentage of elements.
		/// </summary>
		/// <param name="percent">The percentage of elements from the top to show.</param>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = FilterDispatchIds.Filter_TopPercent)]
		void ApplyTopPercentFilter(int percent);

		/// <summary>
		/// Apply a "Icon" filter to the column for the given icon.
		/// </summary>
		/// <param name="icon">The icons of the cells to show.</param>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = FilterDispatchIds.Filter_Icon)]
		void ApplyIconFilter(Icon icon);

		/// <summary>
		/// Apply a "Icon" filter to the column for the given criteria strings.
		/// </summary>
		/// <param name="criteria1">The first criteria string.</param>
		/// <param name="criteria2">The second criteria string.</param>
		/// <param name="oper">The operator that describes how the two criteria are joined.</param>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = FilterDispatchIds.Filter_Custom)]
		void ApplyCustomFilter(string criteria1, [Optional]string criteria2, [Optional]FilterOperator oper);

		/// <summary>
		/// Clear the filter on the given column.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = FilterDispatchIds.Filter_Clear)]
		void Clear();

		/// <summary>
		/// The currently applied filter on the given column.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = FilterDispatchIds.Filter_Criteria)]
		[JsonStringify()]
		FilterCriteria Criteria { get; }
	}

	/// <summary>
	/// Represents the filtering criteria applied to a column.
	/// </summary>
	[ApiSet(Version = 1.2)]
	[ClientCallableComType(Name = "IFilterCriteria", InterfaceId = "BB994BE3-DDDA-4497-9906-8D22855491E8", CoClassName = "FilterCriteria", CoClassId = "9B468C31-9E41-467C-9E5F-24F20B7CB729")]
	public struct FilterCriteria
	{
		/// <summary>
		/// The first criterion used to filter data. Used as an operator in the case of "custom" filtering.
		/// For example ">50" for number greater than 50 or "=*s" for values ending in "s".
		///
		/// Used as a number in the case of top/bottom items/percents. E.g. "5" for the top 5 items if filterOn is set to "topItems"
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = FilterDispatchIds.FilterCriteria_Criterion1)]
		[Optional]
		string Criterion1 { get; set; }

		/// <summary>
		/// The second criterion used to filter data. Only used as an operator in the case of "custom" filtering.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = FilterDispatchIds.FilterCriteria_Criterion2)]
		[Optional]
		string Criterion2 { get; set; }

		/// <summary>
		/// The HTML color string used to filter cells. Used with "cellColor" and "fontColor" filtering. 
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = FilterDispatchIds.FilterCriteria_Color)]
		[Optional]
		string Color { get; set; }

		/// <summary>
		/// The operator used to combine criterion 1 and 2 when using "custom" filtering.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = FilterDispatchIds.FilterCriteria_Operator)]
		[Optional]
		FilterOperator Operator { get; set; }

		/// <summary>
		/// The icon used to filter cells. Used with "icon" filtering.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = FilterDispatchIds.FilterCriteria_Icon)]
		[Optional]
		Icon Icon { get; set; }

		/// <summary>
		/// The dynamic criteria from the Excel.DynamicFilterCriteria set to apply on this column. Used with "dynamic" filtering.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = FilterDispatchIds.FilterCriteria_DynamicCriteria)]
		[Optional]
		DynamicFilterCriteria DynamicCriteria { get; set; }

		/// <summary>
		/// The set of values to be used as part of "values" filtering.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = FilterDispatchIds.FilterCriteria_Values)]
		[Optional]
		[TypeScriptType("Array<string|Excel.FilterDatetime>")]
		[KnownType(typeof(FilterDatetime))]
		object[] Values { get; set; }

		/// <summary>
		/// The property used by the filter to determine whether the values should stay visible.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = FilterDispatchIds.FilterCriteria_FilterOn)]
		FilterOn FilterOn { get; set; }
	}

	/// <summary>
	/// Represents how to filter a date when filtering on values.
	/// </summary>
	[ApiSet(Version = 1.2)]
	[ClientCallableComType(Name = "IFilterDatetime", InterfaceId = "2F73B0F2-4627-4622-8494-640AF24FB44B", CoClassName = "FilterDatetime", CoClassId = "FFAB0D93-6B73-4F3F-8974-4932D35736E2")]
	public struct FilterDatetime
	{
		/// <summary>
		/// The date in ISO8601 format used to filter data.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = FilterDispatchIds.FilterDatetime_Date)]
		string Date { get; set; }

		/// <summary>
		/// How specific the date should be used to keep data. For example, if the date is 2005-04-02 and the specifity is set to "month", the filter operation will keep all rows with a date in the month of april 2009.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = FilterDispatchIds.FilterDatetime_Specificity)]
		FilterDatetimeSpecificity Specificity { get; set; }
	}

	#endregion Filter

#region Images
	internal static class ImagesDispatchIds
	{
		internal const int Icon_Set = 1;
		internal const int Icon_Index = 2;
	}

	/// <summary>
	/// Represents a cell icon.
	/// </summary>
	[ApiSet(Version = 1.2)]
	[ClientCallableComType(Name = "IIcon", InterfaceId = "4FFBA2EE-8527-449C-9C81-739E2795182E", CoClassName = "Icon", CoClassId = "BB897B2C-9B30-4FCA-96B1-E7FFC576FC48")]
	public struct Icon
	{
		/// <summary>
		/// Represents the set that the icon is part of.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = ImagesDispatchIds.Icon_Set)]
		IconSet Set { get; set; }

		/// <summary>
		/// Represents the index of the icon in the given set.
		/// </summary>
		[ApiSet(Version = 1.2)]
		[ClientCallableComMember(DispatchId = ImagesDispatchIds.Icon_Index)]
		int Index { get; set; }
	}
#endregion Images

#region Custom XML Parts
	internal static class CustomXmlDispatchIds
	{
		internal const int CustomXmlPart_OnAccess = 1;
		internal const int CustomXmlPart_Delete = 2;
		internal const int CustomXmlPart_Id = 3;
		internal const int CustomXmlPart_NamespaceUri = 4;
		internal const int CustomXmlPart_GetXml = 5;
		internal const int CustomXmlPart_SetXml = 6;
		internal const int CustomXmlPart_InsertElement = 7;
		internal const int CustomXmlPart_UpdateElement = 8;
		internal const int CustomXmlPart_DeleteElement = 9;
		internal const int CustomXmlPart_Query = 10;
		internal const int CustomXmlPart_InsertAttribute = 11;
		internal const int CustomXmlPart_UpdateAttribute = 12;
		internal const int CustomXmlPart_DeleteAttribute = 13;

		internal const int CustomXmlPartCollection_OnAccess = 1;
		internal const int CustomXmlPartCollection_Indexer = 2;
		internal const int CustomXmlPartCollection_Add = 3;
		internal const int CustomXmlPartCollection_GetByNamespace = 4;
		internal const int CustomXmlPartCollection_GetCount = 5;
		internal const int CustomXmlPartCollection_GetItemOrNullObject = 6;

		internal const int CustomXmlPartScopedCollection_OnAccess = 1;
		internal const int CustomXmlPartScopedCollection_Indexer = 2;
		internal const int CustomXmlPartScopedCollection_GetCount = 3;
		internal const int CustomXmlPartScopedCollection_GetItemOrNullObject = 4;
		internal const int CustomXmlPartScopedCollection_GetOnlyItem = 5;
		internal const int CustomXmlPartScopedCollection_GetOnlyItemOrNullObject = 6;
	}

	/// <summary>
	/// A scoped collection of custom XML parts.
	/// A scoped collection is the result of some operation, e.g. filtering by namespace.
	/// A scoped collection cannot be scoped any further.
	/// </summary>
	[ApiSet(Version = 1.5)]
	[ClientCallableType(UseItemAsIndexerNameInODataId = true)]
	[ClientCallableComType(Name = "ICustomXmlPartScopedCollection", InterfaceId = "2C27E984-EF91-4F48-9A4C-BC96DEF777CE", CoClassName = "CustomXmlPartScopedCollection", SupportEnumeration = true)]
	public interface CustomXmlPartScopedCollection : IEnumerable<CustomXmlPart>
	{
		[ClientCallableComMember(DispatchId = CustomXmlDispatchIds.CustomXmlPartScopedCollection_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Gets a custom XML part based on its ID.
		/// </summary>
		/// <param name="id">ID of the object to be retrieved.</param>
		[ApiSet(Version = 1.5)]
		[ClientCallableComMember(DispatchId = CustomXmlDispatchIds.CustomXmlPartScopedCollection_Indexer)]
		CustomXmlPart this[string id] { get; }

		/// <summary>
		/// Gets the number of CustomXML parts in this collection.
		/// </summary>
		[ApiSet(Version = 1.5)]
		[ClientCallableComMember(DispatchId = CustomXmlDispatchIds.CustomXmlPartScopedCollection_GetCount)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		int GetCount();

		/// <summary>
		/// Gets a custom XML part based on its ID.
		/// If the CustomXmlPart does not exist, the return object's isNull property will be true.
		/// </summary>
		/// <param name="id">ID of the object to be retrieved.</param>
		[ApiSet(Version = 1.5)]
		[ClientCallableComMember(DispatchId = CustomXmlDispatchIds.CustomXmlPartScopedCollection_GetItemOrNullObject)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "")]
		CustomXmlPart GetItemOrNullObject(string id);

		/// <summary>
		/// If the collection contains exactly one item, this method returns it.
		/// Otherwise, this method produces an error.
		/// </summary>
		[ApiSet(Version = 1.5)]
		[ClientCallableComMember(DispatchId = CustomXmlDispatchIds.CustomXmlPartScopedCollection_GetOnlyItem)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		CustomXmlPart GetOnlyItem();

		/// <summary>
		/// If the collection contains exactly one item, this method returns it.
		/// Otherwise, this method returns Null.
		/// </summary>
		[ApiSet(Version = 1.5)]
		[ClientCallableComMember(DispatchId = CustomXmlDispatchIds.CustomXmlPartScopedCollection_GetOnlyItemOrNullObject)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "")]
		CustomXmlPart GetOnlyItemOrNullObject();
	}

	/// <summary>
	/// A collection of custom XML parts.
	/// </summary>
	[ApiSet(Version = 1.5)]
	[ClientCallableComType(Name = "ICustomXmlPartCollection", InterfaceId = "BD3EE512-94FF-4981-9C3A-18F379FAEE41", CoClassName = "CustomXmlPartCollection", SupportEnumeration = true)]
	public interface CustomXmlPartCollection : IEnumerable<CustomXmlPart>
	{
		[ClientCallableComMember(DispatchId = CustomXmlDispatchIds.CustomXmlPartCollection_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Gets a custom XML part based on its ID.
		/// </summary>
		/// <param name="id">ID of the object to be retrieved.</param>
		[ApiSet(Version = 1.5)]
		[ClientCallableComMember(DispatchId = CustomXmlDispatchIds.CustomXmlPartCollection_Indexer)]
		CustomXmlPart this[string id] { get; }

		/// <summary>
		/// Adds a new custom XML part to the workbook.
		/// </summary>
		/// <param name="xml">XML content. Must be a valid XML fragment.</param>
		[ApiSet(Version = 1.5)]
		[ClientCallableComMember(DispatchId = CustomXmlDispatchIds.CustomXmlPartCollection_Add)]
		CustomXmlPart Add(string xml);

		/// <summary>
		/// Gets a new scoped collection of custom XML parts whose namespaces match the given namespace.
		/// </summary>
		/// <param name="namespaceUri"></param>
		[ApiSet(Version = 1.5)]
		[ClientCallableComMember(DispatchId = CustomXmlDispatchIds.CustomXmlPartCollection_GetByNamespace)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		CustomXmlPartScopedCollection GetByNamespace(string namespaceUri);

		/// <summary>
		/// Gets the number of CustomXml parts in the collection.
		/// </summary>
		[ApiSet(Version = 1.5)]
		[ClientCallableComMember(DispatchId = CustomXmlDispatchIds.CustomXmlPartCollection_GetCount)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		int GetCount();

		/// <summary>
		/// Gets a custom XML part based on its ID.
		/// If the CustomXmlPart does not exist, the return object's isNull property will be true.
		/// </summary>
		/// <param name="id">ID of the object to be retrieved.</param>
		[ApiSet(Version = 1.5)]
		[ClientCallableComMember(DispatchId = CustomXmlDispatchIds.CustomXmlPartCollection_GetItemOrNullObject)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "")]
		CustomXmlPart GetItemOrNullObject(string id);
	}

	/// <summary>
	/// Represents a custom XML part object in a workbook.
	/// </summary>
	[ApiSet(Version = 1.5)]
	[ClientCallableType(DeleteOperationName = "Delete")]
	[ClientCallableComType(Name = "ICustomXmlPart", InterfaceId = "2694275E-2EA7-40C8-B98A-EF84C5E22580", CoClassName = "CustomXmlPart")]
	public interface CustomXmlPart
	{
		[ClientCallableComMember(DispatchId = CustomXmlDispatchIds.CustomXmlPart_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Deletes the custom XML part.
		/// </summary>
		[ClientCallableComMember(DispatchId = CustomXmlDispatchIds.CustomXmlPart_Delete)]
		[ApiSet(Version = 1.5)]
		void Delete();

		/// <summary>
		/// The custom XML part's ID. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = CustomXmlDispatchIds.CustomXmlPart_Id)]
		[ApiSet(Version = 1.5)]
		string Id { get; }

		/// <summary>
		/// The custom XML part's namespace URI. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = CustomXmlDispatchIds.CustomXmlPart_NamespaceUri)]
		[ApiSet(Version = 1.5)]
		string NamespaceUri { get; }

		/// <summary>
		/// Gets the custom XML part's full XML content.
		/// </summary>
		[ClientCallableComMember(DispatchId = CustomXmlDispatchIds.CustomXmlPart_GetXml)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		[ApiSet(Version = 1.5)]
		string GetXml();

		/// <summary>
		/// Sets the custom XML part's full XML content.
		/// </summary>
		/// <param name="xml">XML content for the part.</param>
		[ClientCallableComMember(DispatchId = CustomXmlDispatchIds.CustomXmlPart_SetXml)]
		[ApiSet(Version = 1.5)]
		void SetXml(string xml);

	}
	#endregion

	#region PivotTable
	internal static class PivotTableDispatchIds
	{
		internal const int PivotTable_OnAccess = 1;
		internal const int PivotTable_Name = 2;
		internal const int PivotTable_Refresh = 3;
		internal const int PivotTable_Worksheet = 4;
		internal const int PivotTable_Id = 5;

		internal const int PivotTableCollection_OnAccess = 1;
		internal const int PivotTableCollection_Indexer = 2;
		internal const int PivotTableCollection_GetItemOrNullObject = 3;
		internal const int PivotTableCollection_RefreshAll = 4;
		internal const int PivotTableCollection_GetCount = 5;
	}

	/// <summary>
	/// Represents a collection of all the PivotTables that are part of the workbook or worksheet.
	/// </summary>
	[ApiSet(Version = 1.3)]
	[ClientCallableComType(Name = "IPivotTableCollection", InterfaceId = "96495551-83E1-4F20-8B30-EEF756BB1F8D", CoClassName = "PivotTableCollection", SupportEnumeration = true, ExtensibleObject = true)]
	public interface PivotTableCollection : IEnumerable<PivotTable>
	{
		[ClientCallableComMember(DispatchId = PivotTableDispatchIds.PivotTableCollection_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Gets a PivotTable by name.
		/// </summary>
		/// <param name="name">Name of the PivotTable to be retrieved.</param>
		[ApiSet(Version = 1.3)]
		[ClientCallableComMember(DispatchId = PivotTableDispatchIds.PivotTableCollection_Indexer)]

		PivotTable this[string name] { get; }
		/// <summary>
		/// Gets a PivotTable by name. If the PivotTable does not exist, will return a null object.
		/// </summary>
		/// <param name="name">Name of the PivotTable to be retrieved.</param>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = PivotTableDispatchIds.PivotTableCollection_GetItemOrNullObject)]
		[ClientCallableOperation(OperationType = OperationType.Read, RESTfulName = "")]

		PivotTable GetItemOrNullObject(string name);
		/// <summary>
		/// Refreshes all the pivot tables in the collection.
		/// </summary>
		[ApiSet(Version = 1.3)]
		[ClientCallableComMember(DispatchId = PivotTableDispatchIds.PivotTableCollection_RefreshAll)]
		void RefreshAll();

		/// <summary>
		/// Gets the number of pivot tables in the collection.
		/// </summary>
		[ApiSet(Version = 1.4)]
		[ClientCallableComMember(DispatchId = PivotTableDispatchIds.PivotTableCollection_GetCount)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		int GetCount();
	}

	/// <summary>
	/// Represents an Excel PivotTable.
	/// </summary>
	[ApiSet(Version = 1.3)]
	[ClientCallableComType(Name = "IPivotTable", InterfaceId = "1A57CB0A-F84A-4618-B0CC-75CB240CE106", CoClassName = "PivotTable", ExtensibleObject = true)]
	public interface PivotTable
	{
		[ClientCallableComMember(DispatchId = PivotTableDispatchIds.PivotTable_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Id of the PivotTable.
		/// </summary>
		[ApiSet(Version = 1.5)]
		[ClientCallableComMember(DispatchId = PivotTableDispatchIds.PivotTable_Id)]
		string Id { get; }

		/// <summary>
		/// Name of the PivotTable.
		/// </summary>
		[ApiSet(Version = 1.3)]
		[ClientCallableComMember(DispatchId = PivotTableDispatchIds.PivotTable_Name)]
		string Name { get; set; }
		/// <summary>
		/// Refreshes the PivotTable.
		/// </summary>
		[ApiSet(Version = 1.3)]
		[ClientCallableComMember(DispatchId = PivotTableDispatchIds.PivotTable_Refresh)]
		void Refresh();
		/// <summary>
		/// The worksheet containing the current PivotTable. Read-only.
		/// </summary>
		[ApiSet(Version = 1.3)]
		[ClientCallableComMember(DispatchId = PivotTableDispatchIds.PivotTable_Worksheet)]
		Worksheet Worksheet { get; }
	}
#endregion PivotTable

#region Conditional Formats
	internal static class ConditionalFormatDispatchIds
	{
		internal const int ConditionalFormat_Range = 1;
		internal const int ConditionalFormat_Reverse = 2;
		internal const int ConditionalFormat_StopIfTrue = 3;
		internal const int ConditionalFormat_Priority = 4;
		internal const int ConditionalFormat_Type = 5;
		internal const int ConditionalFormat_DataBarOrNullObject = 6;
		internal const int ConditionalFormat_DataBar = 7;
		internal const int ConditionalFormat_CustomOrNullObject = 8;
		internal const int ConditionalFormat_Custom = 9;
		internal const int ConditionalFormat_Delete = 10;
		internal const int ConditionalFormat_OnAccess = 11;
		internal const int ConditionalFormat_RangeOrNull = 12;
		internal const int ConditionalFormat_Icon = 13;
		internal const int ConditionalFormat_IconOrNullObject = 14;
		internal const int ConditionalFormat_ColorScale = 15;
		internal const int ConditionalFormat_ColorScaleOrNullObject = 16;
		internal const int ConditionalFormat_TopBottom = 17;
		internal const int ConditionalFormat_TopBottomOrNullObject = 18;
		internal const int ConditionalFormat_PresetCriteria = 19;
		internal const int ConditionalFormat_PresetCriteriaOrNullObject = 20;
		internal const int ConditionalFormat_Text = 21;
		internal const int ConditionalFormat_TextOrNullObject = 22;
		internal const int ConditionalFormat_CellValue = 23;
		internal const int ConditionalFormat_CellValueOrNullObject = 24;

		internal const int ConditionalFormatCollection_GetCount = 1;
		internal const int ConditionalFormatCollection_ItemAt = 2;
		internal const int ConditionalFormatCollection_ClearAll = 3;
		internal const int ConditionalFormatCollection_Add = 4;
		internal const int ConditionalFormatCollection_OnAccess = 5;

		internal const int ConditionalFormatDataBar_ShowDataBarOnly = 1;
		internal const int ConditionalFormatDataBar_BarDirection = 2;
		internal const int ConditionalFormatDataBar_BorderColor = 3;
		internal const int ConditionalFormatDataBar_AxisFormat = 4;
		internal const int ConditionalFormatDataBar_AxisColor = 5;
		internal const int ConditionalFormatDataBar_PositiveFormat = 6;
		internal const int ConditionalFormatDataBar_NegativeFormat = 7;
		internal const int ConditionalFormatDataBar_LowerBoundRule = 8;
		internal const int ConditionalFormatDataBar_UpperBoundRule = 9;
		internal const int ConditionalFormatDataBar_OnAccess = 10;

		internal const int ConditionalFormatDataBarPositiveFormat_Color = 1;
		internal const int ConditionalFormatDataBarPositiveFormat_IsGradient = 2;
		internal const int ConditionalFormatDataBarPositiveFormat_BorderColor = 3;
		internal const int ConditionalFormatDataBarPositiveFormat_OnAccess = 4;

		internal const int ConditionalFormatDataBarNegativeFormat_Color = 1;
		internal const int ConditionalFormatDataBarNegativeFormat_IsSameColor = 2;
		internal const int ConditionalFormatDataBarNegativeFormat_BorderColor = 3;
		internal const int ConditionalFormatDataBarNegativeFormat_IsSameBorderColor = 4;
		internal const int ConditionalFormatDataBarNegativeFormat_OnAccess = 5;

		internal const int ConditionalFormatDataBarRule_Type = 1;
		internal const int ConditionalFormatDataBarRule_Formula = 2;
		internal const int ConditionalFormatDataBarRule_FormulaLocal = 3;
		internal const int ConditionalFormatDataBarRule_FormulaR1C1 = 4;

		internal const int ConditionalRangeBorder_SideIndex = 1;
		internal const int ConditionalRangeBorder_LineStyle = 2;
		internal const int ConditionalRangeBorder_Color = 3;
		internal const int ConditionalRangeBorder_OnAccess = 4;
		internal const int ConditionalRangeBorder_Id = 5;

		internal const int ConditionalRangeBorderCollection_Indexer = 1;
		internal const int ConditionalRangeBorderCollection_Count = 2;
		internal const int ConditionalRangeBorderCollection_ItemAt = 3;
		internal const int ConditionalRangeBorderCollection_OnAccess = 4;
		internal const int ConditionalRangeBorderCollection_Top = 5;
		internal const int ConditionalRangeBorderCollection_Bottom = 6;
		internal const int ConditionalRangeBorderCollection_Left = 7;
		internal const int ConditionalRangeBorderCollection_Right = 8;

		internal const int ConditionalRangeFill_Color = 1;
		internal const int ConditionalRangeFill_Clear = 2;
		internal const int ConditionalRangeFill_OnAccess = 3;

		internal const int ConditionalRangeFont_Color = 1;
		internal const int ConditionalRangeFont_Italic = 2;
		internal const int ConditionalRangeFont_Bold = 3;
		internal const int ConditionalRangeFont_Underline = 4;
		internal const int ConditionalRangeFont_OnAccess = 5;
		internal const int ConditionalRangeFont_Strikethrough = 6;
		internal const int ConditionalRangeFont_Clear = 7;

		internal const int ConditionalRangeFormat_Fill = 1;
		internal const int ConditionalRangeFormat_Font = 2;
		internal const int ConditionalRangeFormat_Borders = 3;
		internal const int ConditionalRangeFormat_OnAccess = 4;
		internal const int ConditionalRangeFormat_NumberFormat = 5;

		internal const int ConditionalFormatRule_Formula = 1;
		internal const int ConditionalFormatRule_FormulaLocal = 2;
		internal const int ConditionalFormatRule_FormulaR1C1 = 3;
		internal const int ConditionalFormatRule_OnAccess = 4;

		internal const int ConditionalFormatCustom_Rule = 1;
		internal const int ConditionalFormatCustom_OnAccess = 2;
		internal const int ConditionalFormatCustom_Format = 3;

		internal const int ConditionalFormatIcon_ReverseIconOrder = 1;
		internal const int ConditionalFormatIcon_ShowIconOnly = 2;
		internal const int ConditionalFormatIcon_Style = 3;
		internal const int ConditionalFormatIcon_Criterion = 4;
		internal const int ConditionalFormatIcon_OnAccess = 5;

		internal const int ConditionalFormatIconCriterion_Type = 1;
		internal const int ConditionalFormatIconCriterion_Formula = 2;
		internal const int ConditionalFormatIconCriterion_Operator = 3;
		internal const int ConditionalFormatIconCriterion_CustomIcon = 4;

		internal const int ConditionalFormatColorScale_OnAccess = 1;
		internal const int ConditionalFormatColorScale_ThreeColorScale = 2;
		internal const int ConditionalFormatColorScale_Criteria = 3;

		internal const int ConditionalFormatColorScaleCriteria_Minimum = 1;
		internal const int ConditionalFormatColorScaleCriteria_Midpoint = 2;
		internal const int ConditionalFormatColorScaleCriteria_Maximum = 3;

		internal const int ConditionalFormatColorScaleCriterion_Type = 1;
		internal const int ConditionalFormatColorScaleCriterion_Formula = 2;
		internal const int ConditionalFormatColorScaleCriterion_Color = 3;

		internal const int ConditionalFormatTopBottom_OnAccess = 1;
		internal const int ConditionalFormatTopBottom_Rule = 2;
		internal const int ConditionalFormatTopBottom_Format = 3;

		internal const int ConditionalFormatTopBottomRule_Criteria = 1;
		internal const int ConditionalFormatTopBottomRule_Rank = 2;

		internal const int ConditionalFormatPreset_OnAccess = 1;
		internal const int ConditionalFormatPreset_Rule = 2;
		internal const int ConditionalFormatPreset_Format = 3;

		internal const int ConditionalFormatPresetRule_Criteria = 1;

		internal const int ConditionalFormatText_OnAccess = 1;
		internal const int ConditionalFormatText_Rule = 2;
		internal const int ConditionalFormatText_Format = 3;

		internal const int ConditionalFormatTextRule_Operator = 1;
		internal const int ConditionalFormatTextRule_Text = 2;

		internal const int ConditionalFormatCellValue_OnAccess = 1;
		internal const int ConditionalFormatCellValue_Format = 2;
		internal const int ConditionalFormatCellValue_Rule = 3;

		internal const int ConditionalFormatCellValueRule_Operator = 1;
		internal const int ConditionalFormatCellValueRule_Formula1 = 2;
		internal const int ConditionalFormatCellValueRule_Formula2 = 3;
	}

	/// <summary>
	/// Represents a collection of all the conditional formats that are overlap the range.
	/// </summary>
	[ApiSet(Version = 1.6)]
	[ClientCallableComType(Name = "IConditionalFormatCollection", InterfaceId = "34AF8E2C-34B7-4D06-9FF0-08DEB81C7F44", CoClassName = "ConditionalFormatCollection", SupportEnumeration = true)]
	public interface ConditionalFormatCollection : IEnumerable<ConditionalFormat>
	{
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatCollection_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Returns a conditional format at the given index.
		/// </summary>
		/// <param name="index">Index of the conditional formats to be retrieved.</param>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatCollection_ItemAt)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		ConditionalFormat GetItemAt(int index);

		/// <summary>
		/// Returns the number of conditional formats in the workbook. Read-only.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatCollection_GetCount)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		int GetCount();

		/// <summary>
		/// Clears all conditional formats active on the current specified range.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatCollection_ClearAll)]
		void ClearAll();

		/// <summary>
		/// Adds a new conditional format to the collection at the first/top priority.
		/// </summary>
		/// <param name="type">The type of conditional format being added. See Excel.ConditionalFormatType for details.</param>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatCollection_Add)]
		ConditionalFormat Add(ConditionalFormatType type);
	}

	/// <summary>
	/// An object encapsulating a conditional format's range, format, rule, and other properties.
	/// </summary>
	[ClientCallableComType(Name = "IConditionalFormat", InterfaceId = "FED46BE7-0681-4176-A45B-2053C49BC9A8", CoClassName = "ConditionalFormat")]
	[ApiSet(Version = 1.6)]
	public interface ConditionalFormat
	{
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormat_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Returns the range the conditonal format is applied to or a null object if the range is discontiguous. Read-only.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormat_Range)]
		[ClientCallableOperation(OperationType = OperationType.Read, InvalidateReturnObjectPathAfterRequest = true)]
		Range GetRange();

		/// <summary>
		/// Returns the range the conditonal format is applied to or a null object if the range is discontiguous. Read-only.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormat_RangeOrNull)]
		[ClientCallableOperation(OperationType = OperationType.Read, InvalidateReturnObjectPathAfterRequest = true, RESTfulName = "")]
		Range GetRangeOrNullObject();

		/// <summary>
		/// If the conditions of this conditional format are met, no lower-priority formats shall take effect on that cell.
		/// Null on databars, icon sets, and colorscales as there's no concept of StopIfTrue for these
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormat_StopIfTrue)]
		bool? StopIfTrue { get; set; }

		/// <summary>
		/// The priority (or index) within the conditional format collection that this conditional format currently exists in. Changing this also 
		/// changes other conditional formats' priorities, to allow for a contiguous priority order.
		/// Use a negative priority to begin from the back.
		/// Priorities greater than than bounds will get and set to the maximum (or minimum if negative) priority.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormat_Priority)]
		int Priority { get; set; }

		/// <summary>
		/// A type of conditional format. Only one can be set at a time. Read-Only.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormat_Type)]
		ConditionalFormatType Type { get; }

		/// <summary>
		/// Returns the data bar properties if the current conditional format is a data bar.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormat_DataBar)]
		[JsonStringify()]
		DataBarConditionalFormat DataBar { get; }

		/// <summary>
		/// Returns the data bar properties if the current conditional format is a data bar.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormat_DataBarOrNullObject)]
		[JsonStringify()]
		[ClientCallableProperty(ExcludedFromRest = true)]
		DataBarConditionalFormat DataBarOrNullObject { get; }

		/// <summary>
		/// Returns the custom conditional format properties if the current conditional format is a custom type.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormat_Custom)]
		[JsonStringify()]
		CustomConditionalFormat Custom { get; }

		/// <summary>
		/// Returns the custom conditional format properties if the current conditional format is a custom type.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormat_CustomOrNullObject)]
		[ClientCallableProperty(ExcludedFromRest = true)]
		[JsonStringify()]
		CustomConditionalFormat CustomOrNullObject { get; }

		/// <summary>
		/// Returns the IconSet conditional format properties if the current conditional format is an IconSet type.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormat_Icon)]
		[JsonStringify()]
		IconSetConditionalFormat IconSet { get; }

		/// <summary>
		/// Returns the IconSet conditional format properties if the current conditional format is an IconSet type.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormat_IconOrNullObject)]
		[ClientCallableProperty(ExcludedFromRest = true)]
		[JsonStringify()]
		IconSetConditionalFormat IconSetOrNullObject { get; }

		/// <summary>
		/// Returns the ColorScale conditional format properties if the current conditional format is an ColorScale type.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormat_ColorScale)]
		[JsonStringify()]
		ColorScaleConditionalFormat ColorScale { get; }

		/// <summary>
		/// Returns the ColorScale conditional format properties if the current conditional format is an ColorScale type.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormat_ColorScaleOrNullObject)]
		[ClientCallableProperty(ExcludedFromRest = true)]
		[JsonStringify()]
		ColorScaleConditionalFormat ColorScaleOrNullObject { get; }

		/// <summary>
		/// Returns the Top/Bottom conditional format properties if the current conditional format is an TopBottom type.
		/// For example to format the top 10% or bottom 10 items.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormat_TopBottom)]
		[JsonStringify()]
		TopBottomConditionalFormat TopBottom { get; }

		/// <summary>
		/// Returns the Top/Bottom conditional format properties if the current conditional format is an TopBottom type.
		/// For example to format the top 10% or bottom 10 items.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormat_TopBottomOrNullObject)]
		[ClientCallableProperty(ExcludedFromRest = true)]
		[JsonStringify()]
		TopBottomConditionalFormat TopBottomOrNullObject { get; }

		/// <summary>
		/// Returns the preset criteria conditional format such as above average/below average/unique values/contains blank/nonblank/error/noerror properties.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormat_PresetCriteria)]
		[JsonStringify()]
		PresetCriteriaConditionalFormat Preset { get; }

		/// <summary>
		/// Returns the preset criteria conditional format such as above average/below average/unique values/contains blank/nonblank/error/noerror properties.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormat_PresetCriteriaOrNullObject)]
		[ClientCallableProperty(ExcludedFromRest = true)]
		[JsonStringify()]
		PresetCriteriaConditionalFormat PresetOrNullObject { get; }

		/// <summary>
		/// Returns the specific text conditional format properties if the current conditional format is a text type.
		/// For example to format cells matching the word "Text".
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormat_Text)]
		[JsonStringify()]
		TextConditionalFormat TextComparison { get; }

		/// <summary>
		/// Returns the specific text conditional format properties if the current conditional format is a text type.
		/// For example to format cells matching the word "Text".
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormat_TextOrNullObject)]
		[ClientCallableProperty(ExcludedFromRest = true)]
		[JsonStringify()]
		TextConditionalFormat TextComparisonOrNullObject { get; }

		/// <summary>
		/// Returns the cell value conditional format properties if the current conditional format is a CellValue type.
		/// For example to format all cells between 5 and 10.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormat_CellValue)]
		[JsonStringify()]
		CellValueConditionalFormat CellValue { get; }

		/// <summary>
		/// Returns the cell value conditional format properties if the current conditional format is a CellValue type.
		/// For example to format all cells between 5 and 10.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormat_CellValueOrNullObject)]
		[ClientCallableProperty(ExcludedFromRest = true)]
		[JsonStringify()]
		CellValueConditionalFormat CellValueOrNullObject { get; }

		/// <summary>
		/// Deletes this conditional format.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormat_Delete)]
		void Delete();
	}

	/// <summary>
	/// Represents an Excel Conditional Data Bar Type.
	/// </summary>
	[ClientCallableComType(Name = "IDataBarConditionalFormat", InterfaceId = "3378CAB4-80C2-448B-A896-A3BAC8887923", CoClassName = "DataBarConditionalFormat")]
	[ApiSet(Version = 1.6)]
	public interface DataBarConditionalFormat
	{
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatDataBar_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// If true, hides the values from the cells where the data bar is applied.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatDataBar_ShowDataBarOnly)]
		bool ShowDataBarOnly { get; set; }

		/// <summary>
		/// Representation of how the axis is determined for an Excel data bar.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatDataBar_AxisFormat)]
		ConditionalDataBarAxisFormat AxisFormat { get; set; }

		/// <summary>
		/// HTML color code representing the color of the Axis line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
		/// "" (empty string) if no axis is present or set.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatDataBar_AxisColor)]
		string AxisColor { get; set; }

		/// <summary>
		/// Represents the direction that the data bar graphic should be based on.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatDataBar_BarDirection)]
		ConditionalDataBarDirection BarDirection { get; set; }

		/// <summary>
		/// Representation of all values to the right of the axis in an Excel data bar.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatDataBar_PositiveFormat)]
		[JsonStringify()]
		ConditionalDataBarPositiveFormat PositiveFormat { get; }

		/// <summary>
		/// Representation of all values to the left of the axis in an Excel data bar.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatDataBar_NegativeFormat)]
		[JsonStringify()]
		ConditionalDataBarNegativeFormat NegativeFormat { get; }

		/// <summary>
		/// The rule for what consistutes the lower bound (and how to calculate it, if applicable) for a data bar.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatDataBar_LowerBoundRule)]
		[JsonStringify()]
		ConditionalDataBarRule LowerBoundRule { get; set; }

		/// <summary>
		/// The rule for what constitutes the upper bound (and how to calculate it, if applicable) for a data bar.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatDataBar_UpperBoundRule)]
		[JsonStringify()]
		ConditionalDataBarRule UpperBoundRule { get; set; }
	}

	/// <summary>
	/// Represents a conditional format DataBar Format for the positive side of the data bar.
	/// </summary>
	[ClientCallableComType(Name = "IConditionalDataBarPositiveFormat", InterfaceId = "0702CE16-E69F-45E4-A08A-25C6558957BA", CoClassName = "ConditionalDataBarPositiveFormat")]
	[ApiSet(Version = 1.6)]
	public interface ConditionalDataBarPositiveFormat
	{
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatDataBarPositiveFormat_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
		/// "" (empty string) if no border is present or set.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatDataBarPositiveFormat_BorderColor)]
		string BorderColor { get; set; }

		/// <summary>
		/// HTML color code representing the fill color, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatDataBarPositiveFormat_Color)]
		string FillColor { get; set; }

		/// <summary>
		/// Boolean representation of whether or not the DataBar has a gradient.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatDataBarPositiveFormat_IsGradient)]
		bool GradientFill { get; set; }
	}

	/// <summary>
	/// Represents a conditional format DataBar Format for the negative side of the data bar.
	/// </summary>
	[ClientCallableComType(Name = "IConditionalDataBarNegativeFormat", InterfaceId = "631DC6F5-9973-45AD-829C-5339028C37C3", CoClassName = "ConditionalDataBarNegativeFormat")]
	[ApiSet(Version = 1.6)]
	public interface ConditionalDataBarNegativeFormat
	{
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatDataBarNegativeFormat_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
		/// "Empty String" if no border is present or set.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatDataBarNegativeFormat_BorderColor)]
		string BorderColor { get; set; }

		/// <summary>
		/// Boolean representation of whether or not the negative DataBar has the same border color as the positive DataBar.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatDataBarNegativeFormat_IsSameBorderColor)]
		bool MatchPositiveBorderColor { get; set; }

		/// <summary>
		/// HTML color code representing the fill color, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatDataBarNegativeFormat_Color)]
		string FillColor { get; set; }

		/// <summary>
		/// Boolean representation of whether or not the negative DataBar has the same fill color as the positive DataBar.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatDataBarNegativeFormat_IsSameColor)]
		bool MatchPositiveFillColor { get; set; }
	}

	/// <summary>
	/// Represents a rule-type for a Data Bar.
	/// </summary>
	[ClientCallableComType(Name = "IConditionalDataBarRule", InterfaceId = "DECA24F4-4C74-482A-978A-6CC56137A302", CoClassName = "ConditionalDataBarRule", CoClassId = "4CCC8780-5D06-4DD5-BD5D-834DE3AEC7F6")]
	[ApiSet(Version = 1.6)]
	public struct ConditionalDataBarRule
	{
		/// <summary>
		/// The type of rule for the databar.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatDataBarRule_Type)]
		ConditionalFormatRuleType Type { get; set; }

		/// <summary>
		/// The formula, if required, to evaluate the databar rule on.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatDataBarRule_Formula)]
		[TypeScriptType("string")]
		[Optional]
		object Formula { get; set; }
	}

	/// <summary>
	/// Represents a custom conditional format type.
	/// </summary>
	[ClientCallableComType(Name = "ICustomConditionalFormat", InterfaceId = "593C6A29-E4E4-4C0B-AE80-FA808764AB71", CoClassName = "CustomConditionalFormat")]
	[ApiSet(Version = 1.6)]
	public interface CustomConditionalFormat
	{
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatCustom_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the Rule object on this conditional format.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatCustom_Rule)]
		[JsonStringify()]
		ConditionalFormatRule Rule { get; }

		/// <summary>
		/// Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties. Read-only.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatCustom_Format)]
		[JsonStringify()]
		ConditionalRangeFormat Format { get; }
	}

	/// <summary>
	/// Represents a rule, for all traditional rule/format pairings.
	/// </summary>
	[ClientCallableComType(Name = "IConditionalFormatRule", InterfaceId = "55EF76CF-A73F-465D-9F43-EDEE79B6AF95", CoClassName = "ConditionalFormatRule")]
	[ApiSet(Version = 1.6)]
	public interface ConditionalFormatRule
	{
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatRule_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// The formula, if required, to evaluate the conditional format rule on.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatRule_Formula)]
		[TypeScriptType("string")]
		object Formula { get; set; }

		/// <summary>
		/// The formula, if required, to evaluate the conditional format rule on in the user's language.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatRule_FormulaLocal)]
		[TypeScriptType("string")]
		object FormulaLocal { get; set; }

		/// <summary>
		/// The formula, if required, to evaluate the conditional format rule on in R1C1-style notation.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatRule_FormulaR1C1)]
		[TypeScriptType("string")]
		object FormulaR1C1 { get; set; }
	}

	/// <summary>
	/// Represents an IconSet criteria for conditional formatting.
	/// </summary>
	[ClientCallableComType(Name = "IIconSetConditionalFormat", InterfaceId = "994D36C1-E273-43C3-9FCA-D8915C73DFE0", CoClassName = "IconSetConditionalFormat")]
	[ApiSet(Version = 1.6)]
	public interface IconSetConditionalFormat
	{
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatIcon_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		/// <summary>
		/// If true, reverses the icon orders for the IconSet. Note that this cannot be set if custom icons are used.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatIcon_ReverseIconOrder)]
		bool ReverseIconOrder { get; set; }

		/// <summary>
		/// If true, hides the values and only shows icons.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatIcon_ShowIconOnly)]
		bool ShowIconOnly { get; set; }

		/// <summary>
		/// If set, displays the IconSet option for the conditional format.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatIcon_Style)]
		IconSet Style { get; set; }

		/// <summary>
		/// An array of Criteria and IconSets for the rules and potential custom icons for conditional icons. Note that for the first criterion only the custom icon can be modified, while type, formula and operator will be ignored when set.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatIcon_Criterion)]
		[JsonStringify()]
		ConditionalIconCriterion[] Criteria { get; set; }
	}

	/// <summary>
	/// Represents an Icon Criterion which contains a type, value, an Operator, and an optional custom icon, if not using an iconset.
	/// </summary>
	[ClientCallableComType(Name = "IConditionalIconCriterion", InterfaceId = "64CF970B-7712-4BEC-B86E-77AF4B31C989", CoClassName = "ConditionalIconCriterion", CoClassId = "0BBB2E5F-7309-4B25-BB52-2A32D1CC1823")]
	[ApiSet(Version = 1.6)]
	public struct ConditionalIconCriterion
	{
		/// <summary>
		/// What the icon conditional formula should be based on.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatIconCriterion_Type)]
		ConditionalFormatIconRuleType Type { get; set; }

		/// <summary>
		/// A number or a formula depending on the type.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatIconCriterion_Formula)]
		[TypeScriptType("string")]
		object Formula { get; set; }

		/// <summary>
		/// GreaterThan or GreaterThanOrEqual for each of the rule type for the Icon conditional format.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatIconCriterion_Operator)]
		ConditionalIconCriterionOperator Operator { get; set; }

		/// <summary>
		/// The custom icon for the current criterion if different from the default IconSet, else null will be returned.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatIconCriterion_CustomIcon)]
		[Optional]
		Icon CustomIcon { get; set; }
	}

	/// <summary>
	/// Represents an IconSet criteria for conditional formatting.
	/// </summary>
	[ClientCallableComType(Name = "IColorScaleConditionalFormat", InterfaceId = "53F51B37-F896-4CD8-9884-33EB4D84BC61", CoClassName = "ColorScaleConditionalFormat")]
	[ApiSet(Version = 1.6)]
	public interface ColorScaleConditionalFormat
	{
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatColorScale_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// If true the color scale will have three points (minimum, midpoint, maximum), otherwise it will have two (minimum, maximum).
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatColorScale_ThreeColorScale)]
		bool ThreeColorScale { get; }

		/// <summary>
		/// The criteria of the color scale. Midpoint is optional when using a two point color scale.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatColorScale_Criteria)]
		ConditionalColorScaleCriteria Criteria { get; set; }
	}

	/// <summary>
	/// Represents the criteria of the color scale.
	/// </summary>
	[ClientCallableComType(Name = "IConditionalColorScaleCriteria", InterfaceId = "A65ADE8C-F780-40AD-B906-4DE17DC68084", CoClassName = "ConditionalColorScaleCriteria", CoClassId = "06773272-6E84-4E06-924E-C53CE87C3352")]
	[ApiSet(Version = 1.6)]
	public struct ConditionalColorScaleCriteria
	{
		/// <summary>
		/// The minimum point Color Scale Criterion.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatColorScaleCriteria_Minimum)]
		[JsonStringify()]
		ConditionalColorScaleCriterion Minimum { get; set; }

		/// <summary>
		/// The midpoint Color Scale Criterion if the color scale is a 3-color scale.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatColorScaleCriteria_Midpoint)]
		[JsonStringify()]
		[Optional]
		ConditionalColorScaleCriterion Midpoint { get; set; }

		/// <summary>
		/// The maximum point Color Scale Criterion.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatColorScaleCriteria_Maximum)]
		[JsonStringify()]
		ConditionalColorScaleCriterion Maximum { get; set; }
	}

	/// <summary>
	/// Represents a Color Scale Criterion which contains a type, value and a color.
	/// </summary>
	[ClientCallableComType(Name = "IConditionalColorScaleCriterion", InterfaceId = "A00727E4-EB84-4CA9-9D0E-AEEBAF3B9F4E", CoClassName = "ConditionalColorScaleCriterion", CoClassId = "A62AA145-0310-43F4-944C-89D8A59B7303")]
	[ApiSet(Version = 1.6)]
	public struct ConditionalColorScaleCriterion
	{
		/// <summary>
		/// What the icon conditional formula should be based on.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatColorScaleCriterion_Type)]
		ConditionalFormatColorCriterionType Type { get; set; }

		/// <summary>
		/// A number, a formula, or null (if Type is LowestValue).
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatColorScaleCriterion_Formula)]
		[TypeScriptType("string")]
		[Optional]
		object Formula { get; set; }

		/// <summary>
		/// HTML color code representation of the color scale color. E.g. #FF0000 represents Red.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatColorScaleCriterion_Color)]
		[Optional]
		string Color { get; set; }
	}

	/// <summary>
	/// Represents a Top/Bottom conditional format.
	/// </summary>
	[ClientCallableComType(Name = "ITopBottomConditionalFormat", InterfaceId = "3C5F4D98-3AED-47F3-BDA2-2D5BEBCFFF2C", CoClassName = "TopBottomConditionalFormat")]
	[ApiSet(Version = 1.6)]
	public interface TopBottomConditionalFormat
	{
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatTopBottom_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// The criteria of the Top/Bottom conditional format.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatTopBottom_Rule)]
		ConditionalTopBottomRule Rule { get; set; }

		/// <summary>
		/// Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties. Read-only.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatTopBottom_Format)]
		[JsonStringify()]
		ConditionalRangeFormat Format { get; }
	}

	/// <summary>
	/// Represents the rule of the top/bottom conditional format.
	/// </summary>
	[ClientCallableComType(Name = "IConditionalTopBottomRule", InterfaceId = "876FF0AE-2803-4F1F-8FB0-E15184F1BF6D", CoClassName = "ConditionalTopBottomRule", CoClassId = "1DF7829B-9D8D-40CF-A26A-640D368A67A5")]
	[ApiSet(Version = 1.6)]
	public struct ConditionalTopBottomRule
	{
		/// <summary>
		/// Format values based on the top or bottom rank.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatTopBottomRule_Criteria)]
		ConditionalTopBottomCriterionType Type { get; set; }

		/// <summary>
		/// The rank between 1 and 1000 for numeric ranks or 1 and 100 for percent ranks.
		/// </summary>
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatTopBottomRule_Rank)]
		[ApiSet(Version = 1.6)]
		int Rank { get; set; }
	}

	/// <summary>
	/// Represents the the preset criteria conditional format such as above average/below average/unique values/contains blank/nonblank/error/noerror.
	/// </summary>
	[ClientCallableComType(Name = "IPresetCriteriaConditionalFormat", InterfaceId = "AF50DA3F-D7BB-406E-BF67-4FEDF67F2006", CoClassName = "PresetCriteriaConditionalFormat")]
	[ApiSet(Version = 1.6)]
	public interface PresetCriteriaConditionalFormat
	{
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatPreset_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// The rule of the conditional format.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatPreset_Rule)]
		ConditionalPresetCriteriaRule  Rule { get; set; }

		/// <summary>
		/// Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties. Read-only.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatPreset_Format)]
		[JsonStringify()]
		ConditionalRangeFormat Format { get; }
	}

	/// <summary>
	/// Represents the Preset Criteria Conditional Format Rule
	/// </summary>
	[ClientCallableComType(Name = "IConditionalPresetCriteriaRule", InterfaceId = "BCDC2688-54B9-4530-8F3A-0EC2B6589792", CoClassName = "ConditionalPresetCriteriaRule", CoClassId = "11DD755D-A947-4374-9DA7-16470F0F6B7C")]
	[ApiSet(Version = 1.6)]
	public struct ConditionalPresetCriteriaRule
	{
		/// <summary>
		/// The criterion of the conditional format.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatPresetRule_Criteria)]
		ConditionalFormatPresetCriterion Criterion { get; set; }
	}

	/// <summary>
	/// Represents a specific text conditional format.
	/// </summary>
	[ClientCallableComType(Name = "ITextConditionalFormat", InterfaceId = "3CA62D72-357B-4530-9311-332E960B1F96", CoClassName = "TextConditionalFormat")]
	[ApiSet(Version = 1.6)]
	public interface TextConditionalFormat
	{
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatText_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// The rule of the conditional format.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatText_Rule)]
		ConditionalTextComparisonRule Rule { get; set; }

		/// <summary>
		/// Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties. Read-only.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatText_Format)]
		[JsonStringify()]
		ConditionalRangeFormat Format { get; }
	}

	/// <summary>
	/// Represents a Cell Value Conditional Format Rule
	/// </summary>
	[ClientCallableComType(Name = "IConditionalTextComparisonRule", InterfaceId = "6799F838-7DC4-4D64-A890-F125C802F8DD", CoClassName = "ConditionalTextComparisonRule", CoClassId = "E5BA2981-8CAE-459E-940D-8A5724284203")]
	[ApiSet(Version = 1.6)]
	public struct ConditionalTextComparisonRule
	{
		/// <summary>
		/// The operator of the text conditional format.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatTextRule_Operator)]
		ConditionalTextOperator Operator { get; set; }

		/// <summary>
		/// The Text value of conditional format.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatTextRule_Text)]
		string Text { get; set; }
	}

	/// <summary>
	/// Represents a cell value conditional format.
	/// </summary>
	[ClientCallableComType(Name = "ICellValueConditionalFormat", InterfaceId = "7C6B942D-81C8-46A5-BE5C-B8B51A0ED747", CoClassName = "CellValueConditionalFormat")]
	[ApiSet(Version = 1.6)]
	public interface CellValueConditionalFormat
	{
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatCellValue_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the Rule object on this conditional format.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatCellValue_Rule)]
		ConditionalCellValueRule Rule { get; set; }

		/// <summary>
		/// Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties. Read-only.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatCellValue_Format)]
		[JsonStringify()]
		ConditionalRangeFormat Format { get; }
	}

	/// <summary>
	/// Represents a Cell Value Conditional Format Rule
	/// </summary>
	[ClientCallableComType(Name = "IConditionalCellValueRule", InterfaceId = "921BEFA9-AE08-4DF3-B27D-4B86DBC34DA3", CoClassName = "ConditionalCellValueRule", CoClassId = "06ACBCDE-DB8F-4F59-AAD6-E7C4DB98484C")]
	[ApiSet(Version = 1.6)]
	public struct ConditionalCellValueRule
	{
		/// <summary>
		/// The operator of the text conditional format.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatCellValueRule_Operator)]
		ConditionalCellValueOperator Operator { get; set; }

		/// <summary>
		/// The formula, if required, to evaluate the conditional format rule on.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatCellValueRule_Formula1)]
		[TypeScriptType("string")]
		object Formula1 { get; set; }

		/// <summary>
		/// The formula, if required, to evaluate the conditional format rule on.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalFormatCellValueRule_Formula2)]
		[TypeScriptType("string")]
		[Optional]
		object Formula2 { get; set; }
	}

	/// <summary>
	/// A format object encapsulating the conditional formats range's font, fill, borders, and other properties.
	/// </summary>
	[ClientCallableComType(Name = "IConditionalRangeFormat", InterfaceId = "2501401A-42DC-4123-AA48-06AE9F2AB9EE", CoClassName = "ConditionalRangeFormat")]
	[ApiSet(Version = 1.6)]
	public interface ConditionalRangeFormat
	{
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalRangeFormat_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Returns the fill object defined on the overall conditional format range. Read-only.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalRangeFormat_Fill)]
		ConditionalRangeFill Fill { get; }

		/// <summary>
		/// Collection of border objects that apply to the overall conditional format range. Read-only.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalRangeFormat_Borders)]
		ConditionalRangeBorderCollection Borders { get; }

		/// <summary>
		/// Returns the font object defined on the overall conditional format range. Read-only.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalRangeFormat_Font)]
		ConditionalRangeFont Font { get; }


		/// <summary>
		/// Represents Excel's number format code for the given range. Cleared if null is passed in.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalRangeFormat_NumberFormat)]
		object NumberFormat { get; set; }
	}

	/// <summary>
	/// This object represents the font attributes (font style,, color, etc.) for an object.
	/// </summary>
	[ClientCallableComType(Name = "IConditionalRangeFont", InterfaceId = "2AA0159D-35F2-432B-8DA6-D8C7182F90F1", CoClassName = "ConditionalRangeFont")]
	[ApiSet(Version = 1.6)]
	public interface ConditionalRangeFont
	{
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalRangeFont_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();

		/// <summary>
		/// Represents the bold status of font.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalRangeFont_Bold)]
		bool? Bold { get; set; }

		/// <summary>
		/// HTML color code representation of the text color. E.g. #FF0000 represents Red.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalRangeFont_Color)]
		string Color { get; set; }

		/// <summary>
		/// Represents the italic status of the font.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalRangeFont_Italic)]
		bool? Italic { get; set; }

		/// <summary>
		/// Represents the strikethrough status of the font.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalRangeFont_Strikethrough)]
		bool? Strikethrough { get; set; }

		/// <summary>
		/// Type of underline applied to the font. See Excel.ConditionalRangeFontUnderlineStyle for details.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalRangeFont_Underline)]
		ConditionalRangeFontUnderlineStyle? Underline { get; set; }

		/// <summary>
		/// Resets the font formats.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalRangeFont_Clear)]
		void Clear();
	}

	/// <summary>
	/// Represents the background of a conditional range object.
	/// </summary>
	[ClientCallableComType(Name = "IConditionalRangeFill", InterfaceId = "E2488882-675D-4392-8C80-01D6380406D9", CoClassName = "ConditionalRangeFill")]
	[ApiSet(Version = 1.6)]
	public interface ConditionalRangeFill
	{
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalRangeFill_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		/// <summary>
		/// HTML color code representing the color of the fill, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalRangeFill_Color)]
		string Color { get; set; }
		/// <summary>
		/// Resets the fill.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalRangeFill_Clear)]
		void Clear();
	}

	/// <summary>
	/// Represents the border of an object.
	/// </summary>
	[ClientCallableComType(Name = "IConditionalRangeBorder", InterfaceId = "BE3C21C2-66F4-454D-B1BB-7BBCFBA9604B", CoClassName = "ConditionalRangeBorder")]
	[ApiSet(Version = 1.6)]
	public interface ConditionalRangeBorder
	{
		/// <summary>
		/// Represents border identifier. Read-only.
		/// </summary>
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalRangeBorder_Id)]
		[ClientCallableProperty(ExcludedFromClientLibrary = true)]
		[ApiSet(Version = 1.6)]
		ConditionalRangeBorderIndex Id { get; }

		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalRangeBorder_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		/// <summary>
		/// HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalRangeBorder_Color)]
		string Color { get; set; }
		/// <summary>
		/// One of the constants of line style specifying the line style for the border. See Excel.BorderLineStyle for details.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalRangeBorder_LineStyle)]
		ConditionalRangeBorderLineStyle? Style { get; set; }
		/// <summary>
		/// Constant value that indicates the specific side of the border. See Excel.ConditionalRangeBorderIndex for details. Read-only.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalRangeBorder_SideIndex)]
		ConditionalRangeBorderIndex? SideIndex { get; }
	}

	/// <summary>
	/// Represents the border objects that make up range border.
	/// </summary>
	[ClientCallableComType(Name = "IConditionalRangeBorderCollection", InterfaceId = "E18F283E-6188-48FD-A364-C65973CF228C", CoClassName = "ConditionalRangeBorderCollection")]
	[ApiSet(Version = 1.6)]
	public interface ConditionalRangeBorderCollection : IEnumerable<ConditionalRangeBorder>
	{
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalRangeBorderCollection_OnAccess)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		void _OnAccess();
		/// <summary>
		/// Gets a border object using its name
		/// </summary>
		/// <param name="index">Index value of the border object to be retrieved. See Excel.ConditionalRangeBorderIndex for details.</param>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalRangeBorderCollection_Indexer)]
		ConditionalRangeBorder this[ConditionalRangeBorderIndex index] { get; }
		/// <summary>
		/// Number of border objects in the collection. Read-only.
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalRangeBorderCollection_Count)]
		int Count { get; }
		/// <summary>
		/// Gets a border object using its index
		/// </summary>
		/// <param name="index">Index value of the object to be retrieved. Zero-indexed.</param>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalRangeBorderCollection_ItemAt)]
		[ClientCallableOperation(OperationType = OperationType.Read)]
		ConditionalRangeBorder GetItemAt(int index);

		/// <summary>
		/// Gets the top border
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalRangeBorderCollection_Top)]
		ConditionalRangeBorder Top { get; }

		/// <summary>
		/// Gets the top border
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalRangeBorderCollection_Bottom)]
		ConditionalRangeBorder Bottom { get; }

		/// <summary>
		/// Gets the top border
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalRangeBorderCollection_Left)]
		ConditionalRangeBorder Left { get; }

		/// <summary>
		/// Gets the top border
		/// </summary>
		[ApiSet(Version = 1.6)]
		[ClientCallableComMember(DispatchId = ConditionalFormatDispatchIds.ConditionalRangeBorderCollection_Right)]
		ConditionalRangeBorder Right { get; }
	}
	#endregion Conditional Formats

}
