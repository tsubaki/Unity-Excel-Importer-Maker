using UnityEngine;
using System.Collections;
using UnityEditor;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using System.Collections.Generic;
using System.Text;

public class ExcelImporterMaker : EditorWindow
{
	Vector2 curretScroll = new Vector2 (0, 0);

	void OnGUI ()
	{
		GUILayout.Label ("makeing importer", EditorStyles.boldLabel);
		className = EditorGUILayout.TextField ("class name", className);
		
		if (GUILayout.Button ("create")) {
			ExportEntity ();
			ExportImporter ();
			
			AssetDatabase.ImportAsset (filePath);
			AssetDatabase.Refresh (ImportAssetOptions.ForceUpdate);
			Close ();
		}
		
		curretScroll = EditorGUILayout.BeginScrollView (curretScroll);
		EditorGUILayout.BeginVertical("box");
		foreach (ExcelRowParameter cell in typeList) {
			
			cell.isEnable = EditorGUILayout.BeginToggleGroup ("enable", cell.isEnable);
			GUILayout.BeginHorizontal();
			cell.name = EditorGUILayout.TextField (cell.name);
			cell.type = (ValueType)EditorGUILayout.EnumPopup (cell.type);
			GUILayout.EndHorizontal();
			
			EditorGUILayout.EndToggleGroup ();
		}
		EditorGUILayout.EndVertical();
		EditorGUILayout.EndScrollView ();
		
	}	
	
	private enum ValueType
	{
		BOOL,
		STRING,
		INT,
		DOUBLE,
	}
	
	private string filePath = string.Empty;
	private List<ExcelRowParameter> typeList = new List<ExcelRowParameter> ();
	private string className = string.Empty;
	private string fileName = string.Empty;
	
	[MenuItem("Assets/Create XLS Importer")]
	static void ExportExcelToAssetbundle ()
	{
		foreach (Object obj in Selection.objects) {
			
		
			var window = ScriptableObject.CreateInstance<ExcelImporterMaker> ();
			window.filePath = AssetDatabase.GetAssetPath (obj);
			window.fileName = Path.GetFileNameWithoutExtension (window.filePath);
		
		
			using (FileStream stream = File.Open (window.filePath, FileMode.Open, FileAccess.Read)) {
			
				IWorkbook book = new HSSFWorkbook (stream);
			
				ISheet sheet = book.GetSheetAt (0);
				window.className = "Entity_" + sheet.SheetName;
			
				IRow titleRow = sheet.GetRow (0);
				IRow dataRow = sheet.GetRow (1);
				for (int i=0; i < titleRow.LastCellNum; i++) {
					ExcelRowParameter parser = new ExcelRowParameter ();
					parser.name = titleRow.GetCell (i).StringCellValue;
				
					ICell cell = dataRow.GetCell (i);
				
					if (cell.CellType != CellType.Unknown && cell.CellType != CellType.BLANK) {
						parser.isEnable = true;

						try {
							string sampling = cell.StringCellValue;
							parser.type = ValueType.STRING;
						} catch {
						}
						try {
							double sampling = cell.NumericCellValue;
							parser.type = ValueType.DOUBLE;
						} catch {
						}
						try {
							bool sampling = cell.BooleanCellValue;
							parser.type = ValueType.BOOL;
						} catch {
						}
					}
				
					window.typeList.Add (parser);
				}
			
				window.Show ();
			}
		}
	}
	
	void ExportEntity ()
	{
		string entittyTemplate = File.ReadAllText ("Assets/Terasurware/Editor/EntityTemplate.txt");
		StringBuilder builder = new StringBuilder ();
		foreach (ExcelRowParameter row in typeList) {
			if (row.isEnable) {
				builder.AppendLine ();
				builder.AppendFormat ("		public {0} {1};", row.type.ToString ().ToLower (), row.name);
			}
		}
		
		entittyTemplate = entittyTemplate.Replace ("$Types$", builder.ToString ());
		entittyTemplate = entittyTemplate.Replace ("$ExcelData$", className);
		
		File.WriteAllText ("Assets/Terasurware/Classes/" + className + ".cs", entittyTemplate);
	}
	
	void ExportImporter ()
	{
		string importerTemplate = File.ReadAllText ("Assets/Terasurware/Editor/ExportTemplate.txt");
		
		StringBuilder builder = new StringBuilder ();
		int rowCount = 0;
		string tab = "					";
		foreach (ExcelRowParameter row in typeList) {
			if (row.isEnable) {
					
				builder.AppendLine ();
				switch (row.type) {
				case ValueType.BOOL:
					builder.AppendFormat (tab + "p.{0} = row.GetCell({1}).BooleanCellValue;", row.name, rowCount);
					break;
				case ValueType.DOUBLE:
					builder.AppendFormat (tab + "p.{0} = row.GetCell({1}).NumericCellValue;", row.name, rowCount);
					break;
				case ValueType.INT:
					builder.AppendFormat (tab + "p.{0} = (int)row.GetCell({1}).NumericCellValue;", row.name, rowCount);
					break;
				case ValueType.STRING:
					builder.AppendFormat (tab + "p.{0} = row.GetCell({1}).StringCellValue;", row.name, rowCount);
					break;
				}
			}
			rowCount += 1;
		}

		importerTemplate = importerTemplate.Replace ("$IMPORT_PATH$", filePath);
		importerTemplate = importerTemplate.Replace ("$EXPORT_PATH$", Path.ChangeExtension(filePath, ".asset"));
		importerTemplate = importerTemplate.Replace ("$ExcelData$",  className);
		importerTemplate = importerTemplate.Replace ("$EXPORT_DATA$", builder.ToString ());
		importerTemplate = importerTemplate.Replace ("$ExportTemplate$", fileName + "_importer");
			
		File.WriteAllText ("Assets/Terasurware/Classes/Editor/" + fileName + "_importer.cs", importerTemplate);
	}
	
	private class ExcelRowParameter
	{
		public ValueType type;
		public string name;
		public bool isEnable;
	}
}
