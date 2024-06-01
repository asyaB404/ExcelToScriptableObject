using System.Data;
using System.IO;
using Excel;
using UnityEditor;
using UnityEngine;

public class ExcelToolsGUI : EditorWindow
{
    private static string localExcelPath = "Scripts/Editor/Excel";
    private static string localSobjPath = "Scripts/ScriptableObject";
    private static string localAssetsPath = "ScriptableObject";

    [MenuItem("Window/ExeclToSOBJ")]
    public static void ShowWindow()
    {
        GetWindow<ExcelToolsGUI>("Custom Editor Window");
    }

    private void OnGUI()
    {
        GUILayout.Space(20);
        GUILayout.Label("Execl表转SOBJ", EditorStyles.boldLabel);
        localExcelPath = EditorGUILayout.TextField("填写Execl表的路径", localExcelPath);
        GUILayout.Space(20);
        localSobjPath = EditorGUILayout.TextField("填写生成SOBJ的路径", localSobjPath);
        GUILayout.Space(20);
        localAssetsPath = EditorGUILayout.TextField("填写生成SOBJ对应资源文件的路径", localAssetsPath);
        GUILayout.Space(20);
        if (GUILayout.Button("创建SOBJ"))
        {
            SpawnSOBJ();
        }
        GUILayout.Space(20);
        if (GUILayout.Button("创建SOBJ对应的资源文件"))
        {
            SpawnAsset();
        }
    }

    private void SpawnSOBJ()
    {
        DirectoryInfo dInfo = Directory.CreateDirectory(
            Application.dataPath + "/" + localExcelPath
        );
        FileInfo[] files = dInfo.GetFiles();
        DataTableCollection tableConllection;
        string sobjPath = Application.dataPath + "/" + localSobjPath;

        for (int i = 0; i < files.Length; i++)
        {
            if (files[i].Extension != ".xlsx" && files[i].Extension != ".xls")
                continue;
            using FileStream fs = files[i].Open(FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(fs);
            tableConllection = excelReader.AsDataSet().Tables;
            fs.Close();
            excelReader.Close();
            foreach (DataTable table in tableConllection)
            {
                GenerateSObjClass(table, sobjPath);
            }
        }
        AssetDatabase.Refresh();
    }

    private void SpawnAsset()
    {
        string assestPath = "Assets/" + localAssetsPath;
        DirectoryInfo dInfo = Directory.CreateDirectory(
            Application.dataPath + "/" + localExcelPath
        );
        FileInfo[] files = dInfo.GetFiles();
        DataTableCollection tableConllection;
        for (int i = 0; i < files.Length; i++)
        {
            if (files[i].Extension != ".xlsx" && files[i].Extension != ".xls")
                continue;
            using FileStream fs = files[i].Open(FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(fs);
            tableConllection = excelReader.AsDataSet().Tables;
            fs.Close();
            excelReader.Close();
            foreach (DataTable table in tableConllection)
            {
                GenerateAssest(table);
            }
        }
        AssetDatabase.Refresh();
    }

    private void GenerateSObjClass(DataTable table, string sobjPath)
    {
        //字段名行
        DataRow rowName = table.Rows[0];
        //字段类型行
        DataRow rowType = table.Rows[1];

        if (!Directory.Exists(sobjPath))
            Directory.CreateDirectory(sobjPath);
        //如果我们要生成对应的数据结构类脚本 其实就是通过代码进行字符串拼接 然后存进文件就行了
        string str = "using UnityEngine;\n";
        str += "using System.Data;\n\n";
        str +=
            "[CreateAssetMenu(fileName = \""
            + table.TableName
            + "\", menuName = \"ScriptableObject/"
            + table.TableName
            + "\")]\n";
        str += "public class " + table.TableName + " : ExcelableScriptableObject" + "\n{\n";

        for (int j = 0; j < table.Columns.Count; j++)
        {
            if (rowType[j].ToString().Length >= 6 && rowType[j].ToString()[..5] == "Enum.")
            {
                str +=
                    "    public "
                    + rowType[j].ToString()[5..]
                    + " "
                    + rowName[j].ToString()
                    + ";\n";
            }
            else
            {
                str += "    public " + rowType[j].ToString() + " " + rowName[j].ToString() + ";\n";
            }
        }
        str += "\n    public override void Init(DataRow row)\n    {\n";
        for (int j = 0; j < table.Columns.Count; j++)
        {
            if (rowType[j].ToString() == "string")
            {
                str += "        " + rowName[j].ToString() + " = " + "row[" + j + "].ToString();\n";
            }
            else if (rowType[j].ToString().Length >= 6 && rowType[j].ToString()[..5] == "Enum.")
            {
                str +=
                    "        "
                    + rowName[j].ToString()
                    + " = ("
                    + rowType[j].ToString()[5..]
                    + ")System.Enum.Parse(typeof("
                    + rowType[j].ToString()[5..]
                    + "), row["
                    + j
                    + "].ToString());\n";
            }
            else
            {
                str +=
                    "        "
                    + rowName[j].ToString()
                    + " = "
                    + rowType[j].ToString()
                    + ".Parse(row["
                    + j
                    + "].ToString());\n";
            }
        }
        str += "    }\n";
        str += "}";
        File.WriteAllText(sobjPath + "/" + table.TableName + ".cs", str);
    }

    private void GenerateAssest(DataTable table)
    {
        string assetPath = Application.dataPath + "/" + localAssetsPath;
        if (!Directory.Exists(assetPath))
            Directory.CreateDirectory(assetPath);
        DataRow row;
        for (int i = 3; i < table.Rows.Count; i++)
        {
            row = table.Rows[i];
            ScriptableObject obj = ScriptableObject.CreateInstance(table.TableName);
            ExcelableScriptableObject asset = obj as ExcelableScriptableObject;
            asset.Init(row);
            AssetDatabase.CreateAsset(
                asset,
                "Assets/" + localAssetsPath + "/" + table.TableName + "_" + (i - 2) + ".asset"
            );
            AssetDatabase.SaveAssets();
        }
        AssetDatabase.Refresh();
    }
}
