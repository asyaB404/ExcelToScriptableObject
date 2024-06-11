using System.Data;
using UnityEngine;

public abstract class ExcelableScriptableObject : ScriptableObject
{
    public virtual void Init(DataRow row) { }

    public virtual void Init(object[] objects) { }
}
