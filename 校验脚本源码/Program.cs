// See https://aka.ms/new-console-template for more information
using Newtonsoft.Json;
using OfficeOpenXml;
using System.Net.Http.Json;
using System.Security.Cryptography;
using System.Text.Json.Serialization;
using 武器类型配置表数据类型;
using 武器配置表数据类型;
using static System.Runtime.InteropServices.JavaScript.JSType;


ExcelPackage.License.SetNonCommercialPersonal("武器配置表校验工具");

using (var package =new ExcelPackage("C:\\Users\\39369\\Desktop\\数据配置\\武器配置表.xlsx"))
{ 
    var worksheet1=package.Workbook.Worksheets[0];
    int rowcount1 = worksheet1.Dimension.End.Row;
    var worksheet2=package.Workbook.Worksheets[1];
    int rowcount2=worksheet2.Dimension.End.Row;


    List<WeaponItemDatasError> weaponItemDatasError = new List<WeaponItemDatasError>();
    List<WeaponTypeItemDatasError> weaponTypeItemDatasError = new List<WeaponTypeItemDatasError>();

    Dictionary<int,string> keyValuePairs2 = new Dictionary<int,string>();
    Dictionary<int,string> keyValuePairs1= new Dictionary<int,string>();

    for (int i = 2; i <= rowcount2; i++)
    {
     var cellvalue = worksheet2.Cells[i,1].Value;

        if(cellvalue==null) continue;


        if (keyValuePairs2.ContainsKey(Convert.ToInt32(cellvalue)))
        {
            WeaponTypeItemDatasError weaponTypeItemData = new WeaponTypeItemDatasError();
            weaponTypeItemData.TypeID = worksheet2.Cells[i, 1].GetValue<int>();
            weaponTypeItemData.TypeName = worksheet2.Cells[i, 2].GetValue<string>();
            weaponTypeItemData.Tip = "此数据ID重复";
            weaponTypeItemData.tablename = worksheet2.Name;


            weaponTypeItemDatasError.Add(weaponTypeItemData);
        }
        else
        {
            keyValuePairs2.Add(Convert.ToInt32(cellvalue), worksheet2.Cells[i, 2].GetValue<string>());
        }
    }


    for (int i=2;i<=rowcount1;i++)
    {
        var cellvalue = worksheet1.Cells[i, 1].Value;
        if (cellvalue == null) continue;

        if (keyValuePairs1.ContainsKey(Convert.ToInt32(cellvalue)))
        {
            WeaponItemDatasError weaponItemData=new WeaponItemDatasError();

            weaponItemData.WeaponID = worksheet1.Cells[i,1].GetValue<int>();
            weaponItemData.WeaponName=worksheet1.Cells[i, 2].GetValue<string>();
            weaponItemData.Rarity = worksheet1.Cells[i,3].GetValue<string>();
            weaponItemData.WeaponType = worksheet1.Cells[i, 4].GetValue<int>();
            weaponItemData.AttackPwoer = worksheet1.Cells[i,5].GetValue<float>();
            weaponItemData.Tip="此数据ID重复";
            weaponItemData.tablename = worksheet1.Name;

            weaponItemDatasError.Add(weaponItemData);

        }
        else
        {
            keyValuePairs1.Add(Convert.ToInt32(cellvalue), worksheet2.Cells[i, 2].GetValue<string>());
        }
    }


    if (weaponItemDatasError.Count > 0 || weaponTypeItemDatasError.Count > 0)
    {
        Console.WriteLine("存在异常数据，正在生成异常数据表格");


        using (var package1 = new ExcelPackage())

        {
            var worksheet = package1.Workbook.Worksheets.Add("错误数据表格");
            int currentRow = 1;

            if (weaponItemDatasError.Count > 0)
            {
                foreach(var error in weaponItemDatasError)
                {
                    worksheet.Cells[currentRow, 1].Value = error.WeaponID;
                    worksheet.Cells[currentRow, 2].Value = error.WeaponName;
                    worksheet.Cells[currentRow, 3].Value = error.Rarity;
                    worksheet.Cells[currentRow, 4].Value = error.WeaponType;
                    worksheet.Cells[currentRow, 5].Value = error.AttackPwoer;
                    worksheet.Cells[currentRow, 6].Value = error.Tip;
                    worksheet.Cells[currentRow, 7].Value = error.tablename;

                    currentRow++;
                    
                }

                worksheet.Cells.AutoFitColumns();
            }
            if (weaponTypeItemDatasError.Count > 0)
            {

                foreach(var error in weaponTypeItemDatasError)
                {
                    worksheet.Cells[currentRow, 1].Value =error.TypeID;
                    worksheet.Cells[currentRow, 2].Value = error.TypeName;
                    worksheet.Cells[currentRow, 3].Value = error.Tip;
                    worksheet.Cells[currentRow, 4].Value = error.tablename;
                    worksheet.Column(currentRow).Width = 20;

                    currentRow++;
                }
                worksheet.Cells.AutoFitColumns();
            }

            var fileInfo = new FileInfo("C:\\Users\\39369\\Desktop\\数据配置\\异常数据.xlsx");
            using (FileStream fs = fileInfo.OpenWrite())
            {
                package1.SaveAs(fs);
            }

        }



    }
    else
    {
        List<WeaponItemDatas> weaponItemDatas = new List<WeaponItemDatas>();
        for (int i=2;i<= rowcount1;i++)
        {
            WeaponItemDatas weaponitemdatas= new WeaponItemDatas();

            weaponitemdatas.WeaponID = worksheet1.Cells[i, 1].GetValue<int>();
            weaponitemdatas.WeaponName=worksheet1.Cells[i, 2].GetValue<string>();
            weaponitemdatas.Rarity = worksheet1.Cells[i,3].GetValue<string>();
            weaponitemdatas.WeaponType = worksheet1.Cells[i,4].GetValue<int>();
            weaponitemdatas.AttackPwoer = worksheet1.Cells[i,5].GetValue<float>();

            weaponItemDatas.Add(weaponitemdatas);
        }

        string json1= JsonConvert.SerializeObject(weaponItemDatas, Newtonsoft.Json.Formatting.Indented);
        File.WriteAllText("C:\\Users\\39369\\Desktop\\数据配置\\武器配置表.json", json1);


        List<WeaponTypeItemDatas> weaponTypeItemDatas = new List<WeaponTypeItemDatas>();
        for (int i = 2; i <= rowcount2; i++)
        { 
        WeaponTypeItemDatas weapontypeItemdatas = new WeaponTypeItemDatas();
            weapontypeItemdatas.TypeID = worksheet2.Cells[i, 1].GetValue<int>();
            weapontypeItemdatas.TypeName = worksheet2.Cells[i,2].GetValue<string>();

            weaponTypeItemDatas.Add(weapontypeItemdatas);

        }

        string json2 = JsonConvert.SerializeObject(weaponTypeItemDatas, Newtonsoft.Json.Formatting.Indented);
        File.WriteAllText("C:\\Users\\39369\\Desktop\\数据配置\\武器种类配置表.json", json2);

    }

}

namespace 武器配置表数据类型
{
    internal class WeaponItemDatasError
    { 
    public int WeaponID { get; set; }
        public string WeaponName { get; set; }=string.Empty;
        public string Rarity { get; set; }= string.Empty;
        public int WeaponType { get; set; }

        public float AttackPwoer { get; set; }
        public string Tip { get; set; } = string.Empty;
        public string tablename { get; set; } = string.Empty;
    }

    internal class WeaponItemDatas
    {
        public int WeaponID { get; set; }
        public string WeaponName { get; set; } = string.Empty;
        public string Rarity { get; set; } = string.Empty;
        public int WeaponType { get; set; }

        public float AttackPwoer { get; set; }
    }
}

namespace 武器类型配置表数据类型
{
    internal class WeaponTypeItemDatasError
    { 
        public int TypeID {  get; set; }
        public string TypeName { get; set; }=string.Empty;
        public string Tip { get; set; } = string.Empty;
        public string tablename { get; set; } = string.Empty;
    }

    internal class WeaponTypeItemDatas
    {
        public int TypeID { get; set; }
        public string TypeName { get; set; } = string.Empty;

    }
}

