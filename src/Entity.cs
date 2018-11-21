using System;
using System.Reflection;
using System.IO;
using UnityEngine;

[Serializable]
public class Entity
{
    public int ID;


    public virtual void Decode(BinaryReader br){
        br.ReadSingle();
        var a = br.ReadByte() == 0 ? true : false;
    }

    public virtual void PostDecode(){

    }

    
}