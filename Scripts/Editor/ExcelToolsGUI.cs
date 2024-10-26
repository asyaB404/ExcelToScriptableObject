using System;
using System.Data;
using System.IO;
using System.Text;
using Excel;
using UnityEditor;
using UnityEngine;

namespace Basya
{
    public class ExcelToolsGUI : EditorWindow
    {
        private string localExcelPath = "Scripts/Editor/Excel";
        private string localSobjPath = "Scripts/ScriptableObject";
        private string localAssetsPath = "ScriptableObject";

        private string localExcelPath1 = "Scripts/Editor/Excel";
        private string localSobjPath1 = "Scripts/ScriptableObject";
        private string localAssetsPath1 = "ScriptableObject";
        private string localInfoClassPath = "ScriptableObject/InfoClass";

        private readonly StringBuilder strBuilder = new();

        private Panel currentPanel = Panel.Main;

        private Panel CurrentPanel
        {
            get { return currentPanel; }
            set
            {
                currentPanel = value;
                if ((int)currentPanel >= 2)
                    currentPanel = 0;
            }
        }

        private enum Panel
        {
            Main,
            Secondary
        }

        [MenuItem("Window/ExcelToSOBJ")]
        public static void ShowWindow()
        {
            GetWindow<ExcelToolsGUI>("ExcelToSOBJ");
        }

        private void OnGUI()
        {
            GUILayout.Label("Excel表转SOBJ(By Basya)", EditorStyles.boldLabel);
            if (GUILayout.Button("切换页面"))
            {
                CurrentPanel++;
            }

            switch (CurrentPanel)
            {
                case Panel.Main:
                    DrawMainPanel();
                    break;
                case Panel.Secondary:
                    DrawSecondaryPanel();
                    break;
            }
        }

        private void DrawMainPanel()
        {
            GUILayout.Space(20);
            GUILayout.Label("根据Excel表创建的SOBJ来创建多个资源文件", EditorStyles.boldLabel);
            GUILayout.Label("(比较适合如敌人等较复杂的配置信息)", EditorStyles.boldLabel);
            GUILayout.Space(20);
            GUILayout.BeginHorizontal();
            localExcelPath = EditorGUILayout.TextField("Excel表所在文件夹的路径", localExcelPath);
            if (GUILayout.Button("浏览", GUILayout.Width(50)))
            {
                string path = EditorUtility
                    .OpenFolderPanel("选择文件夹", "Assets", "")
                    .Replace(Application.dataPath + "/", "");
                if (!string.IsNullOrEmpty(path))
                {
                    localExcelPath = path;
                }
            }

            GUILayout.EndHorizontal();
            GUILayout.Space(20);
            GUILayout.BeginHorizontal();
            localSobjPath = EditorGUILayout.TextField("填写生成SOBJ的路径", localSobjPath);
            if (GUILayout.Button("浏览", GUILayout.Width(50)))
            {
                string path = EditorUtility
                    .OpenFolderPanel("选择文件夹", "Assets", "")
                    .Replace(Application.dataPath + "/", "");
                if (!string.IsNullOrEmpty(path))
                {
                    localSobjPath = path;
                }
            }

            GUILayout.EndHorizontal();
            GUILayout.Space(20);
            GUILayout.BeginHorizontal();
            localAssetsPath = EditorGUILayout.TextField("生成资源文件的路径", localAssetsPath);
            if (GUILayout.Button("浏览", GUILayout.Width(50)))
            {
                string path = EditorUtility
                    .OpenFolderPanel("选择文件夹", "Assets", "")
                    .Replace(Application.dataPath + "/", "");
                if (!string.IsNullOrEmpty(path))
                {
                    localAssetsPath = path;
                }
            }

            GUILayout.EndHorizontal();
            GUILayout.Space(20);
            if (GUILayout.Button("创建SOBJ"))
            {
                SpawnSOBJ();
            }

            GUILayout.Space(10);
            GUILayout.Label("创建完成后请等待编译结束后", EditorStyles.boldLabel);
            GUILayout.Label("再创建SOBJ对应的资源文件", EditorStyles.boldLabel);
            GUILayout.Space(10);
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
                var file = files[i];
                if (file.Extension != ".xlsx" && file.Extension != ".xls")
                    continue;
                using FileStream fs = file.Open(FileMode.Open, FileAccess.Read);
                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(fs);
                tableConllection = excelReader.AsDataSet().Tables;
                fs.Close();
                excelReader.Close();
                string fileName = Path.GetFileNameWithoutExtension(file.Name);
                foreach (DataTable table in tableConllection)
                {
                    GenerateSObjClass(table, sobjPath, fileName);
                }
            }

            AssetDatabase.Refresh();
        }

        private void GenerateSObjClass(DataTable table, string sobjPath, string fileName)
        {
            DataRow rowName = table.Rows[0];
            DataRow rowType = table.Rows[1];
            if (!Directory.Exists(sobjPath))
                Directory.CreateDirectory(sobjPath);

            strBuilder.Clear();
            strBuilder.AppendLine("using UnityEngine;");
            strBuilder.AppendLine("#if UNITY_EDITOR");
            strBuilder.AppendLine("using System.Data;");
            strBuilder.AppendLine("using UnityEditor;");
            strBuilder.AppendLine("#endif");
            strBuilder.AppendLine();
            strBuilder.AppendLine($"namespace {fileName} ");
            strBuilder.AppendLine("{");
            strBuilder.AppendLine(
                $"\t[CreateAssetMenu(fileName = \"{table.TableName}\", menuName = \"ScriptableObject/{table.TableName}\")]");
            strBuilder.AppendLine($"\tpublic class {table.TableName} : ExcelableScriptableObject");
            strBuilder.AppendLine("\t{");

            GenerateField(table, rowName, rowType);
            GenerateInitMethod(table, rowName, rowType, true);

            strBuilder.AppendLine("\t}");
            strBuilder.AppendLine("}");

            File.WriteAllText(Path.Combine(sobjPath, $"{table.TableName}.cs"), strBuilder.ToString());
        }

        private void GenerateField(DataTable table, DataRow rowName, DataRow rowType)
        {
            for (int j = 0; j < table.Columns.Count; j++)
            {
                if (rowType[j].ToString().Length >= 6 && rowType[j].ToString()[..5] == "Enum.")
                {
                    strBuilder.AppendLine($"\t\tpublic {rowType[j].ToString()[5..]} {rowName[j]};");
                }
                else
                {
                    strBuilder.AppendLine($"\t\tpublic {rowType[j]} {rowName[j]};");
                }
            }
        }

        private void GenerateInitMethod(DataTable table, DataRow rowName, DataRow rowType, bool isOverride)
        {
            strBuilder.AppendLine();
            strBuilder.AppendLine("#if UNITY_EDITOR");
            strBuilder.AppendLine(isOverride
                ? "\t\tpublic override void Init(DataRow row)"
                : "\t\tpublic void Init(DataRow row)");
            strBuilder.AppendLine("\t\t{");
            for (int j = 0; j < table.Columns.Count; j++)
            {
                switch (rowType[j].ToString())
                {
                    case "":
                        break;
                    case "string":
                        strBuilder.AppendLine($"\t\t\t{rowName[j]} = row[{j}].ToString();");
                        break;
                    case "Sprite":
                        strBuilder.AppendLine(
                            $"\t\t\t{rowName[j]} = AssetDatabase.LoadAssetAtPath<Sprite>(row[{j}].ToString().Trim());");
                        break;
                    case var type when type.StartsWith("Enum."):
                        strBuilder.AppendLine(
                            $"\t\t\t{rowName[j]} = ({type[5..]})System.Enum.Parse(typeof({type[5..]}), row[{j}].ToString());");
                        break;
                    default:
                        strBuilder.AppendLine($"\t\t\t{rowName[j]} = {rowType[j]}.Parse(row[{j}].ToString());");
                        break;
                }
            }

            strBuilder.AppendLine("\t\t}");
            strBuilder.AppendLine("#endif");
        }

        private void SpawnAsset()
        {
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
                    GenerateAsset(table);
                }
            }

            AssetDatabase.Refresh();
        }

        private void GenerateAsset(DataTable table)
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

        private void DrawSecondaryPanel()
        {
            GUILayout.Space(20);
            GUILayout.Label("根据Excel表创建自定义类文件", EditorStyles.boldLabel);
            GUILayout.Label("再将其作为对象存储到单个SOBJ资源文件中的列表中", EditorStyles.boldLabel);
            GUILayout.Label("(比较适合关卡配置等轻量级配置信息)", EditorStyles.boldLabel);
            GUILayout.Space(20);
            GUILayout.BeginHorizontal();
            localExcelPath1 = EditorGUILayout.TextField("Excel表所在文件夹的路径", localExcelPath1);
            if (GUILayout.Button("浏览", GUILayout.Width(50)))
            {
                string path = EditorUtility
                    .OpenFolderPanel("选择文件夹", "Assets", "")
                    .Replace(Application.dataPath + "/", "");
                if (!string.IsNullOrEmpty(path))
                {
                    localExcelPath1 = path;
                }
            }

            GUILayout.EndHorizontal();
            GUILayout.Space(20);
            GUILayout.BeginHorizontal();
            localSobjPath1 = EditorGUILayout.TextField("填写生成SOBJ的路径", localSobjPath1);
            if (GUILayout.Button("浏览", GUILayout.Width(50)))
            {
                string path = EditorUtility
                    .OpenFolderPanel("选择文件夹", "Assets", "")
                    .Replace(Application.dataPath + "/", "");
                if (!string.IsNullOrEmpty(path))
                {
                    localSobjPath1 = path;
                }
            }

            GUILayout.EndHorizontal();
            GUILayout.Space(20);
            GUILayout.BeginHorizontal();
            localInfoClassPath = EditorGUILayout.TextField("生成自定义信息类的路径", localInfoClassPath);
            if (GUILayout.Button("浏览", GUILayout.Width(50)))
            {
                string path = EditorUtility
                    .OpenFolderPanel("选择文件夹", "Assets", "")
                    .Replace(Application.dataPath + "/", "");
                if (!string.IsNullOrEmpty(path))
                {
                    localInfoClassPath = path;
                }
            }

            GUILayout.EndHorizontal();
            GUILayout.Space(20);
            GUILayout.BeginHorizontal();
            localAssetsPath1 = EditorGUILayout.TextField("生成资源文件的路径", localAssetsPath1);
            if (GUILayout.Button("浏览", GUILayout.Width(50)))
            {
                string path = EditorUtility
                    .OpenFolderPanel("选择文件夹", "Assets", "")
                    .Replace(Application.dataPath + "/", "");
                if (!string.IsNullOrEmpty(path))
                {
                    localAssetsPath1 = path;
                }
            }

            GUILayout.EndHorizontal();
            GUILayout.Space(20);
            if (GUILayout.Button("创建SOBJ和自定义信息类"))
            {
                SpawnSOBJAndInfoClass();
            }

            GUILayout.Space(10);
            GUILayout.Label("创建完成后请等待编译结束后", EditorStyles.boldLabel);
            GUILayout.Label("再创建SOBJ对应的资源文件", EditorStyles.boldLabel);
            GUILayout.Space(10);
            if (GUILayout.Button("创建SOBJ对应的资源文件"))
            {
                SpawnAsset1();
            }
        }

        private void SpawnSOBJAndInfoClass()
        {
            DirectoryInfo dInfo = Directory.CreateDirectory(
                Application.dataPath + "/" + localExcelPath1
            );
            FileInfo[] files = dInfo.GetFiles();
            DataTableCollection tableConllection;
            string sobjPath = Application.dataPath + "/" + localSobjPath1;
            string infoclassPath = Application.dataPath + "/" + localInfoClassPath;

            for (int i = 0; i < files.Length; i++)
            {
                var file = files[i];
                if (file.Extension != ".xlsx" && file.Extension != ".xls")
                    continue;
                using FileStream fs = file.Open(FileMode.Open, FileAccess.Read);
                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(fs);
                tableConllection = excelReader.AsDataSet().Tables;
                fs.Close();
                excelReader.Close();
                string fileName = Path.GetFileNameWithoutExtension(file.Name + "List");
                foreach (DataTable table in tableConllection)
                {
                    GenerateSObjClass1(table, sobjPath, fileName);
                    GenerateSObjInfoClass(table, infoclassPath, fileName);
                }
            }

            AssetDatabase.Refresh();
        }

        private void GenerateSObjClass1(DataTable table, string sobjPath, string namespaceName)
        {
            if (!Directory.Exists(sobjPath))
                Directory.CreateDirectory(sobjPath);
            string className = table.TableName + "List";
            strBuilder.Clear();
            strBuilder.AppendLine("using System.Collections.Generic;");
            strBuilder.AppendLine("using UnityEngine;");
            strBuilder.AppendLine();
            strBuilder.AppendLine($"namespace {namespaceName}");
            strBuilder.AppendLine("{");
            strBuilder.AppendLine(
                $"\t[CreateAssetMenu(fileName = \"{className}\", menuName = \"ScriptableObject/{table.TableName}\")]");
            strBuilder.AppendLine($"\tpublic class {className} : ExcelableScriptableObject");
            strBuilder.AppendLine("\t{");
            strBuilder.AppendLine($"\t\tpublic List<{className}InfoClass> list = new();");
            strBuilder.AppendLine();
            strBuilder.AppendLine("\t\tpublic override void Init(object[] objects)");
            strBuilder.AppendLine("\t\t{");
            strBuilder.AppendLine("\t\t\tforeach (var obj in objects)");
            strBuilder.AppendLine("\t\t\t{");
            strBuilder.AppendLine($"\t\t\t\tvar obj1 = obj as {className}InfoClass;");
            strBuilder.AppendLine("\t\t\t\tlist.Add(obj1);");
            strBuilder.AppendLine("\t\t\t}");
            strBuilder.AppendLine("\t\t}");
            strBuilder.AppendLine("\t}");
            strBuilder.AppendLine("}"); // 关闭命名空间
            File.WriteAllText(Path.Combine(sobjPath, $"{className}.cs"), strBuilder.ToString());
            AssetDatabase.Refresh();
        }

        private void GenerateSObjInfoClass(DataTable table, string infoClassPath, string namespaceName)
        {
            DataRow rowName = table.Rows[0];
            DataRow rowType = table.Rows[1];
            string className = table.TableName + "ListInfoClass";

            if (!Directory.Exists(infoClassPath))
                Directory.CreateDirectory(infoClassPath);

            strBuilder.Clear();
            strBuilder.AppendLine("using UnityEngine;");
            strBuilder.AppendLine("#if UNITY_EDITOR");
            strBuilder.AppendLine("using System.Data;");
            strBuilder.AppendLine("using UnityEditor;");
            strBuilder.AppendLine("#endif");
            strBuilder.AppendLine();
            strBuilder.AppendLine($"namespace {namespaceName}");
            strBuilder.AppendLine("{");
            strBuilder.AppendLine("\t[System.Serializable]");
            strBuilder.AppendLine($"\tpublic class {className}");
            strBuilder.AppendLine("\t{");

            GenerateField(table, rowName, rowType);
            GenerateInitMethod(table, rowName, rowType, false);

            strBuilder.AppendLine("\t}");
            strBuilder.AppendLine("}"); // 关闭命名空间
            File.WriteAllText(Path.Combine(infoClassPath, $"{className}.cs"), strBuilder.ToString());
            AssetDatabase.Refresh();
        }

        private void SpawnAsset1()
        {
            DirectoryInfo dInfo = Directory.CreateDirectory(
                Application.dataPath + "/" + localExcelPath1
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
                    GenerateAsset1(table);
                }
            }

            AssetDatabase.Refresh();
        }

        private void GenerateAsset1(DataTable table)
        {
            string assetPath = Application.dataPath + "/" + localAssetsPath1;
            if (!Directory.Exists(assetPath))
                Directory.CreateDirectory(assetPath);

            var listClassName = table.TableName + "List";
            ScriptableObject obj = ScriptableObject.CreateInstance(listClassName);

            string className = listClassName + "InfoClass";
            Type type = Type.GetType(className + ", Assembly-CSharp");
            DataRow row;
            object infoObj;
            object[] objects = new object[table.Rows.Count - 3];
            for (int i = 3; i < table.Rows.Count; i++)
            {
                row = table.Rows[i];
                infoObj = Activator.CreateInstance(type);
                type.GetMethod("Init").Invoke(infoObj, new object[] { row });
                objects[i - 3] = infoObj;
            }

            ExcelableScriptableObject asset = obj as ExcelableScriptableObject;
            asset.Init(objects);
            AssetDatabase.CreateAsset(
                asset,
                "Assets/" + localAssetsPath1 + "/" + listClassName + ".asset"
            );
            AssetDatabase.Refresh();
        }
    }
}
