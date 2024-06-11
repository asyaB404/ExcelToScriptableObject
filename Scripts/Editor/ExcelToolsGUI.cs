using System;
using System.Data;
using System.IO;
using System.Text;
using Excel;
using UnityEditor;
using UnityEngine;
using UnityEngine.Events;

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

        private void GenerateSObjClass(DataTable table, string sobjPath)
        {
            //字段名行
            DataRow rowName = table.Rows[0];
            //字段类型行
            DataRow rowType = table.Rows[1];

            if (!Directory.Exists(sobjPath))
                Directory.CreateDirectory(sobjPath);
            //如果我们要生成对应的数据结构类脚本 其实就是通过代码进行字符串拼接 然后存进文件就行了

            strBuilder.Clear();
            strBuilder.Append("using UnityEngine;\n");
            strBuilder.Append("using System.Data;\n\n");
            strBuilder
                .Append("[CreateAssetMenu(fileName = \"")
                .Append(table.TableName)
                .Append("\", menuName = \"ScriptableObject/")
                .Append(table.TableName)
                .Append("\")]\n");
            strBuilder
                .Append("public class ")
                .Append(table.TableName)
                .Append(" : ExcelableScriptableObject\n{\n");

            for (int j = 0; j < table.Columns.Count; j++)
            {
                if (rowType[j].ToString().Length >= 6 && rowType[j].ToString()[..5] == "Enum.")
                {
                    strBuilder
                        .Append("    public ")
                        .Append(rowType[j].ToString()[5..])
                        .Append(" ")
                        .Append(rowName[j].ToString())
                        .Append(";\n");
                }
                else
                {
                    strBuilder
                        .Append("    public ")
                        .Append(rowType[j].ToString())
                        .Append(" ")
                        .Append(rowName[j].ToString())
                        .Append(";\n");
                }
            }

            strBuilder.Append("\n    public override void Init(DataRow row)\n    {\n");
            for (int j = 0; j < table.Columns.Count; j++)
            {
                if (rowType[j].ToString() == "string")
                {
                    strBuilder
                        .Append("        ")
                        .Append(rowName[j].ToString())
                        .Append(" = row[")
                        .Append(j)
                        .Append("].ToString();\n");
                }
                else if (rowType[j].ToString().Length >= 6 && rowType[j].ToString()[..5] == "Enum.")
                {
                    strBuilder
                        .Append("        ")
                        .Append(rowName[j].ToString())
                        .Append(" = (")
                        .Append(rowType[j].ToString()[5..])
                        .Append(")System.Enum.Parse(typeof(")
                        .Append(rowType[j].ToString()[5..])
                        .Append("), row[")
                        .Append(j)
                        .Append("].ToString());\n");
                }
                else
                {
                    strBuilder
                        .Append("        ")
                        .Append(rowName[j].ToString())
                        .Append(" = ")
                        .Append(rowType[j].ToString())
                        .Append(".Parse(row[")
                        .Append(j)
                        .Append("].ToString());\n");
                }
            }

            strBuilder.Append("    }\n").Append("}");

            File.WriteAllText(sobjPath + "/" + table.TableName + ".cs", strBuilder.ToString());
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
                    GenerateAssest(table);
                }
            }
            AssetDatabase.Refresh();
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
                if (files[i].Extension != ".xlsx" && files[i].Extension != ".xls")
                    continue;
                using FileStream fs = files[i].Open(FileMode.Open, FileAccess.Read);
                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(fs);
                tableConllection = excelReader.AsDataSet().Tables;
                fs.Close();
                excelReader.Close();
                foreach (DataTable table in tableConllection)
                {
                    GenerateSObjClass1(table, sobjPath);
                    GenerateSObjInfoClass(table, infoclassPath);
                }
            }
            AssetDatabase.Refresh();
        }

        private void GenerateSObjClass1(DataTable table, string sobjPath)
        {
            if (!Directory.Exists(sobjPath))
                Directory.CreateDirectory(sobjPath);
            strBuilder.Clear();
            strBuilder.AppendLine("using System.Collections.Generic;");
            strBuilder.AppendLine("using UnityEngine;");
            strBuilder.AppendLine();
            strBuilder.AppendLine(
                "[CreateAssetMenu(fileName = \""
                    + table.TableName
                    + "\", menuName = \"ScriptableObject/"
                    + table.TableName
                    + "\")]"
            );
            strBuilder.AppendLine(
                "public class " + table.TableName + " : ExcelableScriptableObject"
            );
            strBuilder.AppendLine("{");
            strBuilder.AppendLine(
                "\tpublic List<" + table.TableName + "InfoClass" + "> list = new();"
            );
            strBuilder.AppendLine();
            strBuilder.AppendLine("\tpublic override void Init(object[] objects)");
            strBuilder.AppendLine("\t{");
            strBuilder.AppendLine("\t\tforeach (var obj in objects)");
            strBuilder.AppendLine("\t\t{");
            strBuilder.AppendLine("\t\t\tvar obj1 = obj as " + table.TableName + "InfoClass" + ";");
            strBuilder.AppendLine("\t\t\tlist.Add(obj1);");
            strBuilder.AppendLine("\t\t}");
            strBuilder.AppendLine("\t}");
            strBuilder.AppendLine("}");
            File.WriteAllText(sobjPath + "/" + table.TableName + ".cs", strBuilder.ToString());
            AssetDatabase.Refresh();
        }

        private void GenerateSObjInfoClass(DataTable table, string infoclassPath)
        {
            //字段名行
            DataRow rowName = table.Rows[0];
            //字段类型行
            DataRow rowType = table.Rows[1];
            string className = table.TableName + "InfoClass";
            if (!Directory.Exists(infoclassPath))
                Directory.CreateDirectory(infoclassPath);
            strBuilder.Clear();
            strBuilder.AppendLine("using System.Data;");
            strBuilder.AppendLine();
            strBuilder.AppendLine("[System.Serializable]");
            strBuilder.AppendLine("public class " + className);
            strBuilder.AppendLine("{");
            for (int j = 0; j < table.Columns.Count; j++)
            {
                if (rowType[j].ToString().Length >= 6 && rowType[j].ToString()[..5] == "Enum.")
                {
                    strBuilder
                        .Append("    public ")
                        .Append(rowType[j].ToString()[5..])
                        .Append(" ")
                        .Append(rowName[j].ToString())
                        .Append(";\n");
                }
                else
                {
                    strBuilder
                        .Append("    public ")
                        .Append(rowType[j].ToString())
                        .Append(" ")
                        .Append(rowName[j].ToString())
                        .Append(";\n");
                }
            }

            strBuilder.Append("\n    public void Init(DataRow row)\n    {\n");
            for (int j = 0; j < table.Columns.Count; j++)
            {
                if (rowType[j].ToString() == "string")
                {
                    strBuilder
                        .Append("        ")
                        .Append(rowName[j].ToString())
                        .Append(" = row[")
                        .Append(j)
                        .Append("].ToString();\n");
                }
                else if (rowType[j].ToString().Length >= 6 && rowType[j].ToString()[..5] == "Enum.")
                {
                    strBuilder
                        .Append("        ")
                        .Append(rowName[j].ToString())
                        .Append(" = (")
                        .Append(rowType[j].ToString()[5..])
                        .Append(")System.Enum.Parse(typeof(")
                        .Append(rowType[j].ToString()[5..])
                        .Append("), row[")
                        .Append(j)
                        .Append("].ToString());\n");
                }
                else
                {
                    strBuilder
                        .Append("        ")
                        .Append(rowName[j].ToString())
                        .Append(" = ")
                        .Append(rowType[j].ToString())
                        .Append(".Parse(row[")
                        .Append(j)
                        .Append("].ToString());\n");
                }
            }

            strBuilder.Append("    }\n").Append("}");
            File.WriteAllText(infoclassPath + "/" + className + ".cs", strBuilder.ToString());
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
                    GenerateAssest1(table);
                }
            }
            AssetDatabase.Refresh();
        }

        private void GenerateAssest1(DataTable table)
        {
            string assetPath = Application.dataPath + "/" + localAssetsPath1;
            if (!Directory.Exists(assetPath))
                Directory.CreateDirectory(assetPath);

            ScriptableObject obj = ScriptableObject.CreateInstance(table.TableName);

            string className = table.TableName + "InfoClass";
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
                "Assets/" + localAssetsPath1 + "/" + table.TableName + ".asset"
            );
            AssetDatabase.Refresh();
        }
    }
}
