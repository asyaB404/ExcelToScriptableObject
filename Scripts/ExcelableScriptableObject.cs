using System.Data;
using UnityEngine;

public abstract class ExcelableScriptableObject : ScriptableObject
{
    public abstract void Init(DataRow row);
}

