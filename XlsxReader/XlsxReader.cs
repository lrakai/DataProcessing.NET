using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml.Linq;
namespace XlsxReader {
	/// <summary>
	/// XlsxReader can read Xlsx Excel Sheets and is
	/// compatible with Windows Desktop and Windows Store Apps.
	/// 
	/// Author: David Kubelka
	/// 
	/// </summary>
	/// <example>
	/// public void ExampleKV()
	///	{
	///		const string resourcename = @"dkExcel.Test.TestData.Test02.xlsx";
	///		var assembly = GetType().GetTypeInfo().Assembly;
	///		using(var iostream = assembly.GetManifestResourceStream(resourcename))
	///		{
	///			var workbook = new XlsxReader(iostream);
	///			foreach(var sheetName in workbook.WorksheetNames)
	///			{
	///				var sheet = workbook[sheetName];
	///				foreach(var row in sheet.Rows)
	///					foreach(var kv in row)
	///						Logger.LogMessage(@"{0}: {1}", kv.Key, kv.Value);
	///			}
	///		}
	///	}
	///	public void ExampleEntity()
	///	{
	///		const string resourcename = @"dkExcel.Test.TestData.Names.xlsx";
	///		var assembly = GetType().GetTypeInfo().Assembly;
	///		using(var iostream = assembly.GetManifestResourceStream(resourcename))
	///		{
	///			var workbook = new XlsxReader(iostream);
	///			foreach(var sheetName in workbook.WorksheetNames)
	///			{
	///				var sheet = workbook[sheetName];
	///				foreach(var row in sheet.Entities<Name>())
	///						Logger.LogMessage(@"{0}, {1}", row.Lastname, row.Firstname);
	///			}
	///		}
	///	}
	/// </example>
	public class XlsxReader {
		#region Xml Namespace Definitions
		static internal XNamespace xmlns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
		static internal XNamespace xmlnsr = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
		static internal XName xmlnsRow = xmlns + @"row";
		static internal XName xmlnsC = xmlns + @"c";
		static internal XName xmlnsR = xmlns + @"r";
		static internal XName xmlnsT = xmlns + @"t";
		static internal XName xmlnsV = xmlns + @"v";
		static internal XName xmlnsF = xmlns + @"f";
		#endregion
		#region XlsxReader
		/// <summary>
		/// Returns a XlsxReader for the specified IO stream.
		/// </summary>
		/// <param name="xlsx"></param>
		public XlsxReader( System.IO.Stream xlsx ) {
			if( xlsx == null )
				throw new ArgumentException( @"Xlsx stream may not be null", @"xlsx" );
			_xlsx = xlsx;
		}
		System.IO.Stream _xlsx;
		#endregion
		#region ColumnAttribute
		/// <summary>
		/// Specify columnname for this property
		/// </summary>
		[AttributeUsage( AttributeTargets.Property )]
		public class ColumnAttribute : System.Attribute {
			public string Name {
				get;
				set;
			}
			public enum ContentTypes {
				Data,
				Formula
			};
			ContentTypes? _contentType;
			public ContentTypes ContentType {
				get {
					if( _contentType.HasValue )
						return _contentType.Value;
					return ContentTypes.Data;
				}
				set {
					_contentType = value;
				}
			}
			bool? _caseSensitive;
			public bool CaseSensitive {
				get {
					if( _caseSensitive.HasValue )
						return _caseSensitive.Value;
					return true;
				}
				set {
					_caseSensitive = value;
				}
			}
		}
		#endregion
		#region Workbook
		ZipArchive GetWorkbook() {
			try {
				return new ZipArchive( _xlsx, ZipArchiveMode.Read );
			}
			catch( Exception x ) {
				throw new Exception( @"Unable to open Excel document as ZipArchive", x );
			}
		}
		ZipArchive _workbook;
		protected ZipArchive Workbook {
			get {
				return _workbook ?? ( _workbook = GetWorkbook() );
			}
		}
		#endregion
		#region Worksheet
		/// <summary>
		/// Returns the available worksheets in this workbook
		/// </summary>
		public IEnumerable<string> WorksheetIds {
			get {
				return from e in Workbook.Entries
					   where @"xl\worksheets".Equals( Path.GetDirectoryName( e.FullName ) )
					   let name = Path.GetFileNameWithoutExtension( e.FullName )
					   orderby name
					   select name;
			}
		}
		public IEnumerable<string> WorksheetNames {
			get {
				return from e in SheetEntries
					   select e.Value.Name;
			}
		}
		/// <summary>
		/// Opens the specified worksheet
		/// </summary>
		/// <param name="sheetName">
		/// Name of the worksheet to be opened
		/// </param>
		/// <param name="hasHeader">
		/// Determines if first row of the
		/// worksheet containts a header.
		/// </param>
		/// <returns>
		/// The requested worksheet
		/// </returns>
		public Sheet GetWorksheet( string sheetName = null, bool hasHeader = true ) {
			return new Sheet( this, sheetName, hasHeader );
		}
		public Sheet GetWorksheetIfExists( string sheetName = null, bool hasHeader = true ) {
			if( this.SheetEntries.ContainsKey( sheetName ) )
				return new Sheet( this, sheetName, hasHeader );
			return null;
		}
		/// <summary>
		/// Raised if the workbook does not contain the requested worksheet 
		/// </summary>
		public class WorksheetDoesNotExistException : Exception {
			public WorksheetDoesNotExistException( string name )
				: base( string.Format( @"The worksheet named {0} does not exist", name ) ) {
			}
		}
		/// <summary>
		/// Returns a worksheet by name
		/// </summary>
		/// <param name="sheetName"></param>
		/// <returns></returns>
		public Sheet this[ string sheetName ] {
			get {
				return GetWorksheet( sheetName );
			}
		}
		Sheet _firstSheet;
		/// <summary>
		/// Returns the first worksheet of the workbook
		/// </summary>
		public Sheet FirstSheet {
			get {
				return _firstSheet ?? ( _firstSheet = GetWorksheet() );
			}
		}
		/// <summary>
		/// Represents a specific Excel worksheet in an Excel workbook.
		/// </summary>
		public class Sheet {
			XlsxReader _xlsx;
			ZipArchiveEntry _worksheet;
			/// <summary>
			/// Specifies if the first row of the worksheet contains
			/// a header to be used as either propertynames or keynames.
			/// </summary>
			public bool HasHeader {
				get {
					return _hasHeader;
				}
			}
			bool _hasHeader;
			const string rIdPrefix = @"rId";
			IEnumerable<KeyValuePair<string, string>> KeyedValues( XElement row, CellContent.GetValue getValue = null ) {
				int i = 0;
				foreach( var value in _xlsx.CellContents.Values( row, getValue ) ) {
					if( value != null )
						yield return new KeyValuePair<string, string>( i < Header.Length ? Header[ i ] ?? CellAddress.GetDefaultHeader( i ) : CellAddress.GetDefaultHeader( i ), value );
					i += 1;
				}
			}
			static XName RelationshipName = XName.Get( @"Relationship", @"http://schemas.openxmlformats.org/package/2006/relationships" );
			string GetTarget( string id ) {
				var rels = GetDocument( _xlsx.Workbook.GetEntry( @"xl/_rels/workbook.xml.rels" ) );
				return ( from r in rels.Root.Elements( RelationshipName )
						 where r.Attribute( @"Id" ).Value == id
						 select r.Attribute( @"Target" ).Value ).FirstOrDefault();
			}
			public Sheet( XlsxReader xlsx, string sheetName = null, bool hasHeader = true ) {
				_xlsx = xlsx;
				_hasHeader = hasHeader;
				var name = sheetName ?? _xlsx.WorksheetNames.FirstOrDefault() ?? @"sheet1";
				if( !_xlsx.SheetEntries.ContainsKey( name ) )
					throw new WorksheetDoesNotExistException( name );
				var sheetEntry = _xlsx.SheetEntries[ name ];
				_worksheet = _xlsx.Workbook.GetEntry( $"xl/{GetTarget( sheetEntry.RId )}" );
				if( _worksheet == null )
					throw new WorksheetDoesNotExistException( name );
			}
			/// <summary>
			/// Contains the Xml document representing the worksheet
			/// </summary>
			public XDocument Document {
				get {
					return _document ?? ( _document = GetDocument( _worksheet ) );
				}
			}

			XDocument _document;
			XDocument GetDocument( ZipArchiveEntry document ) {
				using( var documentStream = document.Open() )
					return XDocument.Load( documentStream );
			}
			/// <summary>
			/// Contains the Xml sheet data
			/// </summary>
			public XElement SheetData {
				get {
					return _sheetData ?? ( _sheetData = Document.Root.Element( xmlns + @"sheetData" ) );
				}
			}
			XElement _sheetData;
			/// <summary>
			/// Gets or sets the header of a worksheet,
			/// default is the first row of the worksheet
			/// </summary>
			public string[] Header {
				get {
					return _header ?? ( _header = GetHeader() );
				}
				set {
					if( value == null || value.Length == 0 )
						_hasHeader = false;
					_header = value;
				}
			}
			private string[] _header;
			private string[] GetHeader() {
				if( !HasHeader || !SheetData.HasElements )
					return new string[ 0 ];
				if( _headerAddress == null ) {
					var header = this.ElementsOfAllRows.First();
					HeaderAddress = new CellAddress( header.Elements( xmlnsC ).First() );
					return Trim( _xlsx.CellContents.Values( header ) ).ToArray();
				}
				return Trim( _xlsx.CellContents.Values( GetRow( HeaderAddress.RowName ) ) ).ToArray();
			}
			private IEnumerable<string> Trim( IEnumerable<string> values ) {
				foreach( var value in values )
					if( string.IsNullOrEmpty( value ) )
						yield return value;
					else
						yield return value.Trim();
			}
			public CellAddress HeaderAddress {
				get {
					if( _headerAddress == null )
						Header = GetHeader();
					return _headerAddress;
				}
				set {
					_headerAddress = value;
					_hasHeader = _headerAddress != null;
					_header = null;
				}
			}
			CellAddress _headerAddress;
			internal IEnumerable<XElement> ElementsOfAllRows {
				get {
					if( SheetData.HasElements )
						foreach( var row in SheetData.Elements( xmlnsRow ) )
							yield return row;
				}
			}
			protected IEnumerable<XElement> ElementsOfDataRows {
				get {
					if( SheetData.HasElements ) {
						var headerRow = HasHeader ? HeaderAddress.Row : 0;
						foreach( var row in ElementsOfAllRows )
							if( HasHeader ) {
								if( CellAddress.RowFromRowName( row ) > headerRow )
									yield return row;
							}
							else
								yield return row;
					}
				}
			}
			/// <summary>
			/// Read worksheet and return rows consiting of keyvalue pairs
			/// </summary>
			public IEnumerable<IEnumerable<KeyValuePair<string, string>>> Rows {
				get {
					foreach( var row in ElementsOfDataRows )
						yield return KeyedValues( row );
				}
			}
			/// <summary>
			/// Read worksheet and call bool cellAction(zeroBasedColumn,zeroBasedRow,data) for each cell
			/// and continue if true is returned
			/// </summary>
			public void CellsByRow( Func<int, int, string, bool> cellAction ) {
				foreach( var row in ElementsOfAllRows ) {
					int y = CellAddress.RowFromRowName( row );
					int x = 0;
					foreach( var value in _xlsx.CellContents.Values( row ) ) {
						if( value != null )
							if( !cellAction( x, y, value ) )
								return;
						x += 1;
					}
				}
			}
			class ColumnMap : Dictionary<string, ColumnAttribute> {
				Type _t;
				static private StringComparer GetComparer( Type t ) {
					foreach( var p in t.GetRuntimeProperties() ) {
						var column = p.GetCustomAttribute<ColumnAttribute>();
						if( column != null )
							if( !column.CaseSensitive )
								return StringComparer.OrdinalIgnoreCase;
					}
					return StringComparer.Ordinal;
				}
				public ColumnMap( Type t )
					: base( GetComparer( t ) ) {
					_t = t;
					foreach( var p in t.GetRuntimeProperties() ) {
						var column = p.GetCustomAttribute<ColumnAttribute>();
						var columnName = column == null ? p.Name : column.Name ?? p.Name;
						this[ columnName ] = new ColumnAttribute {
							Name = p.Name,
							ContentType = column == null ? ColumnAttribute.ContentTypes.Data : column.ContentType,
							CaseSensitive = column == null ? true : column.CaseSensitive
						};
					}
				}
				public PropertyInfo GetPropertyInfo( string columnName ) {
					if( string.IsNullOrEmpty( columnName ) || !this.ContainsKey( columnName ) )
						return null;
					var c = this[ columnName ];
					return _t.GetRuntimeProperty( c.Name );
				}
				public ColumnAttribute.ContentTypes GetContentType( string columnName ) {
					if( string.IsNullOrEmpty( columnName ) || !this.ContainsKey( columnName ) )
						return ColumnAttribute.ContentTypes.Data;
					return this[ columnName ].ContentType;
				}
			}
			/// <summary>
			/// Read worksheet and return data in class T
			/// with properties set to the values in worksheet.
			/// </summary>
			/// <typeparam name="T">
			/// Class that as properties matching column names. Properties might be decorated with
			/// a CollumnAttribute
			/// </typeparam>
			/// <returns>
			/// Returns instances of type T with the porperties set
			/// that match a column name of the sheet
			/// </returns>
			public IEnumerable<T> Entities<T>() where T : class, new() {
				var t = typeof( T );
				var columnMap = new ColumnMap( t );
				foreach( var row in ElementsOfDataRows ) {
					var entity = new T();
					var hasValue = false;
					foreach( var cell in KeyedValues( row, ( XElement e, int i ) => {
						var name = i < Header.Length ? Header[ i ] : null;
						if( columnMap.GetContentType( name ) == ColumnAttribute.ContentTypes.Data )
							return this._xlsx.CellContents.Value( e );
						return this._xlsx.CellContents.Formula( e );
					} ) ) {
						var destination = columnMap.GetPropertyInfo( cell.Key );
						if( destination == null )
							continue;
						destination.SetValue( entity, Convert.ChangeType( cell.Value, destination.PropertyType ) );
						hasValue = true;
					}
					if( hasValue )
						yield return entity;
				}
			}
			/// <summary>
			/// Returns the value at the specified coordinates in the worksheet
			/// </summary>
			/// <param name="column">
			/// Column zero based, Excel column A == 0
			/// </param>
			/// <param name="row">
			/// Row zero based, Excel row 1 == 0
			/// </param>
			public string this[ int column, int row ] {
				get {
					return this[ CellAddress.GetColumnName( column, row ), CellAddress.GetRowName( row ) ];
				}
			}
			/// <summary>
			/// Returns the value from the cell with the specified cellname
			/// </summary>
			/// <param name="cellName">
			/// Cellname Excel style A1, AA1, BA1, etc...
			/// </param>
			public string this[ string cellName ] {
				get {
					return this[ cellName, CellAddress.GetRowName( cellName ) ];
				}
			}
			private XElement GetRow( string rowName ) {
				return ( from r in ElementsOfAllRows
						 where rowName.Equals( r.Attribute( @"r" ).Value )
						 select r ).FirstOrDefault();
			}
			private XElement GetCell( string colName, XElement row ) {
				if( row == null || !row.HasElements )
					return null;
				return ( from c in row.Elements( xmlnsC )
						 where colName.Equals( c.Attribute( @"r" ).Value )
						 select c ).FirstOrDefault();
			}
			internal XElement GetLeftCorner() {
				if( SheetData.HasElements ) {
					var minOffset = int.MaxValue;
					foreach( var offset in from r in this.ElementsOfAllRows
										   where r.HasElements
										   select CellAddress.ColumnOffset( r.Elements( xmlnsC ).First() ) )
						if( offset < minOffset )
							minOffset = offset;
					return ( from r in this.ElementsOfAllRows
							 where r.HasElements && CellAddress.ColumnOffset( r.Elements( xmlnsC ).First() ) == minOffset
							 select r ).FirstOrDefault();
				}
				return null;
			}
			private string this[ string col, string row ] {
				get {
					return _xlsx.CellContents.Value( GetCell( col, GetRow( row ) ) );
				}
			}
		}
		#endregion
		#region SheetEntries
		class SheetEntry {
			public string Name {
				get;
				set;
			}
			public string Id {
				get;
				set;
			}
			public string RId {
				get;
				set;
			}
		}
		class SheetEntryDictionary : Dictionary<string, SheetEntry> {
			public SheetEntryDictionary( ZipArchive workbook )
				: base( StringComparer.OrdinalIgnoreCase ) {
				foreach( var sheet in ReadSheetEntries( workbook ) )
					this[ sheet.Name ] = sheet;
			}
			private IEnumerable<SheetEntry> ReadSheetEntries( ZipArchive workbook ) {
				var book = workbook.GetEntry( @"xl/workbook.xml" );
				if( book == null )
					return Enumerable.Empty<SheetEntry>();
				using( var stream = book.Open() ) {
					var document = XDocument.Load( stream );
					return from e in document.Root.Element( xmlns + @"sheets" ).Elements( xmlns + @"sheet" )
						   select new SheetEntry {
							   Name = e.Attribute( @"name" ).Value,
							   Id = e.Attribute( @"sheetId" ).Value,
							   RId = e.Attribute( xmlnsr + @"id" ).Value
						   };
				}
			}
		}
		SheetEntryDictionary _sheetEntries;
		private SheetEntryDictionary SheetEntries {
			get {
				return _sheetEntries ?? ( _sheetEntries = new SheetEntryDictionary( Workbook ) );
			}
			set {
				_sheetEntries = value;
			}
		}
		#endregion
		#region CellAddress
		public class CellAddress {
			const char RadixBase = 'A';
			const char RadixLast = 'Z';
			const int Radix = 1 + RadixLast - RadixBase;
			static int Power( int y, int x ) {
				var value = 1;
				for( int i = 0; i < x; i++ )
					value *= y;
				return value;
			}
			static int LetterCount( string value ) {
				var i = 0;
				foreach( var c in value )
					if( char.IsLetter( c ) )
						i += 1;
					else
						break;
				return i;
			}
			static IEnumerable<char> Letters( string value ) {
				foreach( var c in value )
					if( char.IsLetter( c ) )
						yield return c;
					else
						break;
			}
			static IEnumerable<char> Digits( string value ) {
				foreach( var c in value )
					if( char.IsDigit( c ) )
						yield return c;
			}
			internal static int RowOffset( string r ) {
				var b = new StringBuilder( r.Length );
				foreach( var c in Digits( r ) )
					b.Append( c );
				return int.Parse( b.ToString() ) - 1;
			}
			internal static int RowOffset( XElement row ) {
				return int.Parse( row.Attribute( @"r" ).Value ) - 1;
			}
			public static int ColumnOffset( string r ) {
				var digits = new Stack<int>();
				foreach( var c in Letters( r ) )
					digits.Push( c - RadixBase );
				if( digits.Count < 2 )
					return digits.Pop();
				int value = 0;
				int d = 0;
				while( digits.Count > 0 )
					value += Power( Radix, d++ ) * digits.Pop();
				return Radix + value;
			}
			public static int ColumnOffset( XElement cell ) {
				return ColumnOffset( cell.Attribute( @"r" ).Value );
			}
			public static string GetRowName( int row ) {
				return ( row + 1 ).ToString();
			}
			public static string GetRowName( string cell ) {
				return cell.Substring( LetterCount( cell ) );
			}
			public static string GetColumnName( int column, int row ) {
				var name = new System.Text.StringBuilder();
				if( column < Radix ) // different system for 'A' - 'Z'!
					name.Append( (char) ( RadixBase + column ) );
				else {
					var digits = new Stack<int>( Digits( column - Radix, Radix ) );
					if( digits.Count < 2 ) // need leading zero?
						name.Append( RadixBase );
					while( digits.Count > 0 )
						name.Append( (char) ( RadixBase + digits.Pop() ) );
				}
				name.Append( row + 1 );
				return name.ToString();
			}
			static IEnumerable<int> Digits( int number, int radix ) {
				var v = number;
				do {
					yield return v % radix;
					v /= radix;
				} while( v > 0 );
			}
			string _sheet;
			string _address;
			int? _row;
			int? _column;
			public static readonly CellAddress Origin = new CellAddress( @"A1" );
			public CellAddress( CellAddress value )
				: this( value.ToString() ) {
			}
			public CellAddress( XElement cell )
				: this( cell.Attribute( @"r" ).Value ) {
			}
			public CellAddress( string address ) {
				try {
					if( string.IsNullOrEmpty( address ) ) {
						_sheet = null;
						_address = null;
						_row = null;
						_column = null;
					}
					else {
						var s = address.Split( '!' );
						if( s.Length == 2 ) {
							_sheet = s[ 0 ].Trim( '\'' );
							_address = s[ 1 ];
						}
						else {
							_sheet = null;
							_address = s[ 0 ];
						}
						_row = RowOffset( _address );
						_column = ColumnOffset( _address );
					}
				}
				catch {
					throw new Exception( string.Format( @"Invalid cell address: >>{0}<<", address ) );
				}
			}
			public override string ToString() {
				if( string.IsNullOrEmpty( Sheet ) )
					return this.Address;
				return string.Format( @"{0}!{1}", Sheet, Address );
			}
			public string Sheet {
				get {
					return _sheet;
				}
			}
			public string Address {
				get {
					return _address;
				}
			}
			public int Column {
				get {
					if( _column.HasValue )
						return _column.Value;
					throw new Exception( @"Cell's address has now column" );
				}
				set {
					_column = value;
					_address = _row.HasValue ? GetColumnName( _column.Value, _row.Value ) : null;
				}
			}
			public int Row {
				get {
					if( _row.HasValue )
						return _row.Value;
					throw new Exception( @"Cell's address has no row" );
				}
				set {
					_row = value;
					_address = _column.HasValue ? GetColumnName( _column.Value, _row.Value ) : null;
				}
			}
			public string RowName {
				get {
					return ( _row + 1 ).ToString();
				}
			}
			public static int RowFromRowName( XElement row ) {
				return int.Parse( row.Attribute( @"r" ).Value ) - 1;
			}
			internal static string GetDefaultHeader( int column ) {
				var name = new System.Text.StringBuilder();
				name.Append( '?' );
				if( column < Radix ) // different system for 'A' - 'Z'!
					name.Append( (char) ( RadixBase + column ) );
				else {
					var digits = new Stack<int>( Digits( column - Radix, Radix ) );
					if( digits.Count < 2 ) // need leading zero?
						name.Append( RadixBase );
					while( digits.Count > 0 )
						name.Append( (char) ( RadixBase + digits.Pop() ) );
				}
				return name.ToString();
			}
		}
		#endregion
		#region CellContent
		CellContent _cellContents;
		public CellContent CellContents {
			get {
				return _cellContents ?? ( _cellContents = new CellContent( Workbook ) );
			}
			set {
				_cellContents = value;
			}
		}
		public class CellContent {
			ZipArchive Workbook {
				get;
				set;
			}
			public CellContent( ZipArchive workbook ) {
				Workbook = workbook;
			}
			internal class SharedString {
				public string[] Strings {
					get;
					set;
				}
				public string this[ int offset ] {
					get {
						return Strings[ offset ];
					}
				}
				public SharedString( ZipArchive workbook ) {
					Strings = ReadSharedStrings( workbook ).ToArray();
				}
				private IEnumerable<string> ReadSharedStrings( ZipArchive workbook ) {
					var sheet = workbook.GetEntry( @"xl/sharedStrings.xml" );
					if( sheet == null )
						return Enumerable.Empty<string>();
					using( var stream = sheet.Open() ) {
						var document = XDocument.Load( stream );
						return from e in document.Root.Elements( xmlns + @"si" )
							   select StringValue( e );
					}
				}
				public static string StringValue( XElement si ) {
					var t = si.Elements( xmlnsT ).FirstOrDefault();
					if( t != null )
						return t.Value;
					var b = new StringBuilder();
					foreach( var r in si.Elements( xmlnsR ) )
						b.Append( r.Element( xmlnsT ).Value );
					return b.ToString();
				}
			}
			SharedString _sharedStrings;
			internal SharedString SharedStrings {
				get {
					return _sharedStrings ?? ( _sharedStrings = new SharedString( Workbook ) );
				}
			}
			public string Value( XElement cell ) {
				if( cell == null )
					return null;
				var element = cell.Elements( xmlnsV ).FirstOrDefault();
				if( element == null )
					return null;
				var value = element.Value;
				var type = cell.Attribute( @"t" );
				if( type == null )
					return value;
				switch( type.Value ) {
					case @"s":
						return SharedStrings[ Convert.ToInt32( value ) ];
					case @"str":
						return value;
					default:
						return value;
				}
			}
			public string Formula( XElement cell ) {
				var element = cell.Elements( xmlnsF ).FirstOrDefault();
				if( element == null )
					return null;
				return element.Value;
			}
			internal delegate string GetValue( XElement cell, int offset );
			internal IEnumerable<string> Values( XElement row, GetValue getValue = null ) {
				if( row != null ) {
					int col = 0;
					foreach( var cell in row.Elements( xmlnsC ) ) {
						var nextColumn = CellAddress.ColumnOffset( cell );
						while( nextColumn > col ) {
							yield return null;
							col += 1;
						}
						if( getValue == null )
							yield return Value( cell );
						else
							yield return getValue( cell, col );
						col += 1;
					}
				}
			}
		}
		#endregion
	}
}
