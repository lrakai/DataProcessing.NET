using System.Linq;
using System.Collections.Generic;
using System.Xml.Linq;
namespace XlsxReader {
	/// <summary>
	/// XlsxPivotReader can unpivot Xlsx Excel Sheets and is
	/// compatible with Windows Desktop and Windows Store Apps.
	/// 
	/// Author: David Kubelka
	/// 
	/// </summary>
	public class XlsxPivotReader : XlsxReader {
		#region Constructors
		public XlsxPivotReader( System.IO.Stream xlsx )
			: base( xlsx ) {
		}
		#endregion
		#region Dimensions
		private Dimension[] Dimensions {
			get {
				return _dimensions ?? ( _dimensions = GetDimensions() );
			}
			set {
				_dimensions = value;
			}
		}
		Dimension[] _dimensions;
		Dimension[] GetDimensions() {
			return Dimension.Enumerate( this ).ToArray();
		}
		public class Dimension {
			[Column( Name = @"Key", CaseSensitive = false )]
			public string Name {
				get;
				set;
			}
			public enum Orientations {
				Vertical,
				Horizontal
			}
			public Orientations Orientation {
				get;
				set;
			}
			public int Rank {
				get;
				set;
			}
			public string Sheet {
				get;
				set;
			}
			public static IEnumerable<Dimension> Enumerate( XlsxPivotReader w ) {
				var multiTable = w.MultiTableWorkbook;
				var sheet = w.GetWorksheetIfExists( SheetNaming.PivotHorizontal );
				if( sheet != null ) {
					var r = 0;
					foreach( var d in sheet.Entities<Dimension>() ) {
						d.Rank = r++;
						d.Orientation = Dimension.Orientations.Horizontal;
						yield return d;
					}
				}
				sheet = w.GetWorksheetIfExists( SheetNaming.PivotVertical );
				if( sheet != null ) {
					var r = 0;
					foreach( var d in sheet.Entities<Dimension>() ) {
						d.Rank = r++;
						d.Orientation = Dimension.Orientations.Vertical;
						yield return d;
					}
				}
			}
			public static IEnumerable<Dimension> Enumerate( XlsxPivotReader w, string tableName ) {
				var multiTable = w.MultiTableWorkbook;
				var sheet = w.GetWorksheetIfExists( SheetNaming.PivotHorizontal );
				if( sheet != null ) {
					var r = 0;
					foreach( var d in sheet.Entities<Dimension>() ) {
						if( multiTable )
							if( !tableName.Equals( d.Sheet, System.StringComparison.OrdinalIgnoreCase ) )
								continue;
						d.Rank = r++;
						d.Orientation = Dimension.Orientations.Horizontal;
						yield return d;
					}
				}
				sheet = w.GetWorksheetIfExists( SheetNaming.PivotVertical );
				if( sheet != null ) {
					var r = 0;
					foreach( var d in sheet.Entities<Dimension>() ) {
						if( multiTable )
							if( !tableName.Equals( d.Sheet, System.StringComparison.OrdinalIgnoreCase ) )
								continue;
						d.Rank = r++;
						d.Orientation = Dimension.Orientations.Vertical;
						yield return d;
					}
				}
			}
		}
		#endregion
		#region MetaDataTables
		public MetaDataTable[] MetaDataTables {
			get {
				return _metaDataTables ?? ( _metaDataTables = MetaDataTable.Read( this ).ToArray() );
			}
			set {
				_metaDataTables = value;
			}
		}
		MetaDataTable[] _metaDataTables;
		internal IEnumerable<MetaDataTable> MetaDataTableByName( string _tableName ) {
			if( this.MultiTableWorkbook )
				return from md in MetaDataTables
					   where _tableName.Equals( md.Sheet, System.StringComparison.OrdinalIgnoreCase )
					   select md;
			return from md in MetaDataTables
				   select md;
		}
		public class MetaDataTable {
			[Column( ContentType = ColumnAttribute.ContentTypes.Formula, CaseSensitive = false )]
			public string DataOrigin {
				get;
				set;
			}
			[Column( ContentType = ColumnAttribute.ContentTypes.Formula, CaseSensitive = false )]
			public string LabelOrigin {
				get;
				set;
			}
			[Column( CaseSensitive = false )]
			public string Sheet {
				get {
					return _sheet ?? new CellAddress( LabelOrigin ).Sheet ?? new CellAddress( DataOrigin ).Sheet;
				}
				set {
					_sheet = value;
				}
			}
			string _sheet;
			[Column( CaseSensitive = false )]
			public string Table {
				get;
				set;
			}
			[Column( CaseSensitive = false )]
			public string Measure {
				get {
					return _measure ?? @"Value";
				}
				set {
					_measure = value;
				}
			}
			private string _measure;
			private bool? _deleteBeforeBulkLoad;
			[Column( CaseSensitive = false )]
			public bool DeleteBeforeBulkLoad {
				get {
					if( _deleteBeforeBulkLoad.HasValue )
						return _deleteBeforeBulkLoad.Value;
					return false;
				}
				set {
					_deleteBeforeBulkLoad = value;
				}
			}
			[Column( CaseSensitive = false )]
			public string SqlServer {
				get {
					return _sqlServer ?? @"(local)";
				}
				set {
					_sqlServer = value;
				}
			}
			string _sqlServer;
			[Column( CaseSensitive = false )]
			public string Database {
				get;
				set;
			}
			[Column( CaseSensitive = false )]
			public string ConnectionOptions {
				get;
				set;
			}
			public static IEnumerable<MetaDataTable> Read( XlsxReader xlsx ) {
				var workSheet = xlsx.GetWorksheetIfExists( SheetNaming.PivotSheets );
				if( workSheet != null )
					foreach( var sheet in workSheet.Entities<MetaDataTable>() )
						yield return sheet;
			}
			public static IEnumerable<string> PivotTableSheets( IEnumerable<MetaDataTable> metaDataTables ) {
				return ( ( from o in
							   ( from m in metaDataTables
								 where m.LabelOrigin != null
								 select m.LabelOrigin ).Union( from m in metaDataTables
															   where m.DataOrigin != null
															   select m.DataOrigin )
						   select new CellAddress( o ).Sheet ).Union( from m in metaDataTables
																	  where !string.IsNullOrEmpty( m.Sheet )
																	  select m.Sheet ) ).Distinct();
			}
			public static MetaDataTable Get( IEnumerable<MetaDataTable> metaDataTables ) {
				return metaDataTables.FirstOrDefault() ?? new MetaDataTable();
			}
		}
		#endregion
		#region TableNames
		public string[] TableNames {
			get {
				return _tableNames ?? ( _tableNames = SheetNaming.AllTableSheetNames( this ).ToArray() );
			}
			set {
				_tableNames = value;
			}
		}
		string[] _tableNames;
		public bool MultiTableWorkbook {
			get {
				return TableNames.Length > 1;
			}
		}
		#endregion
		#region Tables
		public IEnumerable<PivotTable> Tables {
			get {
				return from name in TableNames
					   select new PivotTable( this, name );
			}
		}
		#endregion
		#region DefaultTable
		public PivotTable DefaultTable {
			get {
				return _defaultTable ?? ( _defaultTable = Table.GetDefaultTable() );
			}
			set {
				_defaultTable = value;
			}
		}
		PivotTable _defaultTable;
		#endregion
		#region Table [ name ]
		public TableSelector Table {
			get {
				return _table ?? ( _table = TableSelector.GetTableSelector( this ) );
			}
			set {
				_table = value;
			}
		}
		public class TableSelector {
			private XlsxPivotReader _xlsx;
			private TableSelector( XlsxPivotReader xlsx ) {
				_xlsx = xlsx;
			}
			public PivotTable this[ string tableName ] {
				get {
					return new PivotTable( _xlsx, tableName );
				}
			}
			public PivotTable GetDefaultTable( string tableName = null ) {
				if( string.IsNullOrEmpty( tableName ) )
					if( _xlsx.MultiTableWorkbook )
						throw new System.Exception( @"Multiple table present, need to specify worksheet name" );
				return new PivotTable( _xlsx, tableName ?? _xlsx.TableNames.FirstOrDefault() );
			}
			public static TableSelector GetTableSelector( XlsxPivotReader xlsx ) {
				return new TableSelector( xlsx );
			}
		}
		TableSelector _table;
		#endregion
		#region SheetNaming
		class SheetNaming {
			public const string Prefix = @"_pivot";
			public const string PivotSheets = Prefix + @"sheets";
			public const string PivotData = Prefix + @"data";
			public const string PivotHorizontal = Prefix + @"horizontal";
			public const string PivotVertical = Prefix + @"vertical";
			[Column( Name = @"Sheet" )]
			public string Name {
				get;
				set;
			}
			public static IEnumerable<string> AllTableSheetNames( XlsxPivotReader book ) {
				var pivotWorkSheetNames = from n in book.WorksheetNames
										  let name = n.ToLowerInvariant()
										  where
												name.StartsWith( Prefix ) &&
												!PivotSheets.Equals( name ) &&
												!name.StartsWith( PivotData )
										  select n;
				var pivotSheetNames = from sheetName in pivotWorkSheetNames
									  select book[ sheetName ].Entities<SheetNaming>();
				var pivotTableNames = from sheetNames in pivotSheetNames
									  from sheetName in sheetNames
									  select sheetName.Name;
				var allNames = ( from n in pivotTableNames.Union( MetaDataTable.PivotTableSheets( book.MetaDataTables ) )
								 where !string.IsNullOrEmpty( n )
								 select n ).Distinct( System.StringComparer.OrdinalIgnoreCase ).ToList();
				if( allNames.Count > 0 )
					return allNames;
				return from n in book.WorksheetNames
					   where !n.ToLowerInvariant().StartsWith( SheetNaming.Prefix )
					   select n;
			}
		}
		#endregion
		#region PivotTable
		public class PivotTable : IEnumerable<IEnumerable<KeyValuePair<string, string>>> {
			XlsxPivotReader _xlsx;
			public string TableName {
				get {
					return _tableName;
				}
				set {
					_tableName = value;
				}
			}
			string _tableName;
			public MetaDataTable MetaData {
				get {
					return _metaData ?? ( _metaData = _xlsx.MetaDataTableByName( TableName ).FirstOrDefault() ?? new MetaDataTable() );
				}
			}
			MetaDataTable _metaData;
			Dimension[] _dimensions;
			public Dimension[] Dimensions {
				get {
					return _dimensions ?? ( _dimensions = GetTableDimensions().ToArray() );
				}
			}
			Dimension[] _horizontal;
			IEnumerable<Dimension> GetTableDimensions() {
				if( _xlsx.MultiTableWorkbook )
					return from d in _xlsx.Dimensions
						   where TableName.Equals( d.Sheet, System.StringComparison.OrdinalIgnoreCase )
						   select d;
				return from d in _xlsx.Dimensions
					   select d;
			}
			public Dimension[] Horizontal {
				get {
					return _horizontal ?? ( _horizontal = ( from d in Dimensions
															where d.Orientation == Dimension.Orientations.Horizontal
															select d ).ToArray() );
				}
			}
			Dimension[] _vertical;
			public Dimension[] Vertical {
				get {
					return _vertical ?? ( _vertical = ( from d in Dimensions
														where d.Orientation == Dimension.Orientations.Vertical
														select d ).ToArray() );
				}
			}
			public string DatabaseTable {
				get {
					return MetaData.Table ?? TableName;
				}
			}
			public string[] ColumnNames {
				get {
					return _columnNames ?? ( _columnNames = GetColumnNames().ToArray() );
				}
				set {
					_columnNames = value;
					DataSheet.Header = value;
				}
			}
			string[] _columnNames;
			private IEnumerable<string> GetColumnNames() {
				if( Dimensions.Length == 0 ) {
					DataSheet.HeaderAddress = LabelOrigin;
					foreach( var column in DataSheet.Header )
						if( column != null )
							yield return column;
				}
				else {
					yield return MetaData.Measure;
					foreach( var d in Dimensions )
						yield return d.Name;
				}
			}
			CellAddress GetDataOrigin() {
				if( MetaData != null && !string.IsNullOrEmpty( this.MetaData.DataOrigin ) )
					try {
						return new CellAddress( this.MetaData.DataOrigin );
					}
					catch {
					}
				var lo = new CellAddress( LabelOrigin );
				lo.Row += Dimensions.Length == 0 ? 1 : Horizontal.Length;
				lo.Column += Dimensions.Length == 0 ? 0 : Vertical.Length;
				return lo;
			}
			CellAddress _dataOrigin;
			public CellAddress DataOrigin {
				get {
					return _dataOrigin ?? ( _dataOrigin = GetDataOrigin() );
				}
			}
			CellAddress GetLabelOrigin() {
				if( MetaData != null ) {
					if( !string.IsNullOrEmpty( this.MetaData.LabelOrigin ) ) {
						return new CellAddress( this.MetaData.LabelOrigin );
					}
					if( !string.IsNullOrEmpty( MetaData.DataOrigin ) ) {
						var lo = new CellAddress( MetaData.DataOrigin );
						if( Dimensions.Length == 0 ) {
							lo.Row = -1;
						}
						else {
							lo.Row -= Horizontal.Length;
							lo.Column -= Vertical.Length;
						}
						return lo;
					}
				}
				var lc = DataSheet.GetLeftCorner();
				return lc == null ? CellAddress.Origin : new CellAddress( lc.Elements( xmlnsC ).First() );
			}
			CellAddress _labelOrigin;
			public CellAddress LabelOrigin {
				get {
					return _labelOrigin ?? ( _labelOrigin = GetLabelOrigin() );
				}
			}
			public Sheet DataSheet {
				get {
					return _dataSheet ?? ( _dataSheet = GetDataSheet() );
				}
				set {
					_dataSheet = value;
				}
			}
			Sheet _dataSheet;
			bool OriginSpecified {
				get {
					return !( MetaData == null || ( MetaData.DataOrigin == null && MetaData.LabelOrigin == null ) );
				}
			}
			Sheet GetDataSheet() {
				return _xlsx.GetWorksheet( sheetName: OriginSpecified ? LabelOrigin.Sheet : null, hasHeader: Dimensions.Length == 0 ? true : false );
			}
			internal PivotTable( XlsxPivotReader pr, string tableName ) {
				_xlsx = pr;
				DataSheet = _xlsx.GetWorksheet( tableName, false );
				_tableName = tableName;
			}
			class TableEnumerator {
				PivotTable _pt;
				int _maxTableColumn = 0;
				int _maxTableRow = 0;
				List<DimensionValue>[] _horizontalValues;
				List<DimensionValue>[] _verticalValues;
				static List<DimensionValue>[] InitDimensionValues( List<DimensionValue>[] values ) {
					for( int i = 0; i < values.Length; i++ )
						values[ i ] = new List<DimensionValue>();
					return values;
				}
				private List<DimensionValue>[] GetVerticalValues() {
					return InitDimensionValues( new List<DimensionValue>[ _pt.Vertical.Length ] );
				}
				private List<DimensionValue>[] GetHorizontalValues() {
					return InitDimensionValues( new List<DimensionValue>[ _pt.Horizontal.Length ] );
				}
				public TableEnumerator( PivotTable pt ) {
					_pt = pt;
					_horizontalValues = GetHorizontalValues();
					_verticalValues = GetVerticalValues();
				}
				struct DimensionValue {
					public int Offset;
					public string Value;
				}
				IEnumerable<KeyValuePair<string, string>> GetDimensionsVertical( int row ) {
					for( int i = 0; i < _verticalValues.Length; i++ ) {
						DimensionValue? last = null;
						foreach( var dv in _verticalValues[ i ] ) {
							if( dv.Offset > row ) {
								if( last.HasValue )
									yield return new KeyValuePair<string, string>( _pt.Vertical[ i ].Name, last.Value.Value );
								last = null;
								break;
							}
							last = dv;
						}
						if( last.HasValue )
							yield return new KeyValuePair<string, string>( _pt.Vertical[ i ].Name, last.Value.Value );
					}
				}
				IEnumerable<KeyValuePair<string, string>> GetDimensionsHorizontal( int column ) {
					for( int i = 0; i < _horizontalValues.Length; i++ ) {
						DimensionValue? last = null;
						foreach( var dv in _horizontalValues[ i ] ) {
							if( dv.Offset > column ) {
								if( last.HasValue )
									yield return new KeyValuePair<string, string>( _pt.Horizontal[ i ].Name, last.Value.Value );
								last = null;
								break;
							}
							last = dv;
						}
						if( last.HasValue )
							yield return new KeyValuePair<string, string>( _pt.Horizontal[ i ].Name, last.Value.Value );
					}
				}
				KeyValuePair<string, string>? GetMeasure( string value, int column, int row ) {
					if( _pt.DataOrigin.Row > row ) {
						var rowDimension = row - _pt.DataOrigin.Row + _pt.Horizontal.Length;
						if( rowDimension >= 0 && !string.IsNullOrEmpty( value ) ) {
							_horizontalValues[ rowDimension ].Add( new DimensionValue {
								Offset = column,
								Value = value
							} );
							if( column > _maxTableColumn )
								_maxTableColumn = column;
						}
						return null;
					}
					if( _pt.DataOrigin.Column > column ) {
						var colDimension = column - _pt.DataOrigin.Column + _pt.Vertical.Length;
						if( colDimension >= 0 && !string.IsNullOrEmpty( value ) ) {
							for( int i = colDimension; i < _pt.Vertical.Length; i++ )
								_verticalValues[ i ].Clear();
							_verticalValues[ colDimension ].Add( new DimensionValue {
								Offset = row,
								Value = value
							} );
							if( row > _maxTableRow )
								_maxTableRow = row;
						}
						return null;
					}
					if( column > _maxTableColumn )
						_maxTableColumn = column;
					if( row > _maxTableRow )
						_maxTableRow = row;
					return new KeyValuePair<string, string>( _pt.MetaData.Measure, value );
				}
				IEnumerable<KeyValuePair<string, string>> GetDataRow( KeyValuePair<string, string> measure, int column, int row ) {
					yield return measure;
					foreach( var dimension in GetDimensionsHorizontal( column ) )
						yield return dimension;
					foreach( var dimension in GetDimensionsVertical( row ) )
						yield return dimension;
				}
				public IEnumerable<IEnumerable<KeyValuePair<string, string>>> GetDataRows( XElement dataSheetRow ) {
					var rowOffset = CellAddress.RowOffset( dataSheetRow );
					int columnOffset = 0;
					foreach( var cellValue in _pt._xlsx.CellContents.Values( dataSheetRow ) ) {
						var measure = GetMeasure( cellValue, column: columnOffset, row: rowOffset );
						if( measure.HasValue )
							yield return GetDataRow( measure: measure.Value, column: columnOffset, row: rowOffset );
						columnOffset += 1;
					}
					for( int c = columnOffset; c <= _maxTableColumn; c++ ) {
						var measure = GetMeasure( null, column: columnOffset, row: rowOffset );
						if( measure.HasValue )
							yield return GetDataRow( measure: measure.Value, column: c, row: rowOffset );
					}
				}
				internal IEnumerator<IEnumerable<KeyValuePair<string, string>>> GetRows() {
					if( _pt.Dimensions.Length > 0 ) {
						foreach( var dataSheetRow in _pt.DataSheet.ElementsOfAllRows )
							foreach( var pivotTableRow in this.GetDataRows( dataSheetRow ) )
								yield return pivotTableRow;
					}
					else {
						if( _pt._columnNames == null )
							_pt.DataSheet.HeaderAddress = _pt.LabelOrigin;
						foreach( var dataSheetRow in _pt.DataSheet.Rows )
							yield return dataSheetRow;
					}
				}
			}
			IEnumerator<IEnumerable<KeyValuePair<string, string>>> IEnumerable<IEnumerable<KeyValuePair<string, string>>>.GetEnumerator() {
				return new TableEnumerator( this ).GetRows();
			}
			System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() {
				return new TableEnumerator( this ).GetRows();
			}
		}
		#endregion
	}
}
