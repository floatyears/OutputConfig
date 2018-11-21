using System;
using System.Reflection;
using System.IO;
using UnityEngine;

public class Main : MonoBehaviour
{
    void Start(){
        var reader = new BinaryReader(new FileStream("Assets/Build/config/data_test_config.bytes", FileMode.Open));
        var len = reader.ReadInt32();
        for(int i = 0; i < len; i++){
            var cfg = new EntityTestConfig();
            cfg.Decode(reader);
            Debug.Log(cfg);
        }
    }

    void Destroy(){

    }
    
}