using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using Newtonsoft.Json;

public class WeaponDataLoader : MonoBehaviour
{
    // Start is called before the first frame update
    void Start()
    {
        TextAsset jsontext = Resources.Load<TextAsset>("武器配置表");

        if (jsontext == null)
        {
            Debug.LogError("未找到对应文件");
            return;
        }

        List<WeaponItems> weaponlist = JsonConvert.DeserializeObject<List<WeaponItems>>(jsontext.text);
        if (weaponlist == null || weaponlist.Count == 0)
        {
            Debug.LogError("json解析失败或数据为空");
            return;
        }

        for (int i = 0; i <= weaponlist.Count-1; i++)
        {
            Debug.Log($"配置成功，武器名称{weaponlist[i].WeaponName}，伤害为{weaponlist[i].AttackPwoer}" );
        }

    }


}
