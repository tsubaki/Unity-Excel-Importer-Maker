using UnityEngine;
using System.Collections;
using System.IO;
using UnityEditor;
using System.Xml.Serialization;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

public class param1_importer : AssetPostprocessor {
	private static readonly string filePath = "Assets/ExcelData/param1.xls";
	private static readonly string exportPath = "Assets/ExcelData/param1.asset";
	
	static void OnPostprocessAllAssets (string[] importedAssets, string[] deletedAssets, string[] movedAssets, string[] movedFromAssetPaths)
	{
		foreach (string asset in importedAssets) {
			if (!filePath.Equals (asset))
				continue;
				
			Entity_Gofe data = (Entity_Gofe)AssetDatabase.LoadAssetAtPath (exportPath, typeof(Entity_Gofe));
			if (data == null) {
				data = ScriptableObject.CreateInstance<Entity_Gofe> ();
				AssetDatabase.CreateAsset ((ScriptableObject)data, exportPath);
				data.hideFlags = HideFlags.NotEditable;
			}
			
			data.list.Clear ();
			using (FileStream stream = File.Open (filePath, FileMode.Open, FileAccess.Read)) {
				IWorkbook book = new HSSFWorkbook (stream);
				ISheet sheet = book.GetSheetAt (0);
				
				for (int i=1; i< sheet.LastRowNum; i++) {
					IRow row = sheet.GetRow (i);
					Entity_Gofe.Param p = new Entity_Gofe.Param ();
					
					p.skill = row.GetCell(0).StringCellValue;
					p.effect = row.GetCell(1).StringCellValue;
					p.damage = row.GetCell(2).NumericCellValue;
					p.isEnable = row.GetCell(3).BooleanCellValue;
					data.list.Add (p);
				}
			}

			ScriptableObject obj = AssetDatabase.LoadAssetAtPath (exportPath, typeof(ScriptableObject)) as ScriptableObject;
			EditorUtility.SetDirty (obj);
		}
	}
}
