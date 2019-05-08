using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Web;

namespace ExcelDealApp.Models
{
    public class YCLib
    {
        /// <summary>
        /// 加密
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public string ConvertMD5(string data)

        {

            byte[] Original = System.Text.Encoding.Default.GetBytes(data); //將字串來源轉為Byte[]
    
        System.Security.Cryptography.MD5 s1 = System.Security.Cryptography.MD5.Create(); //使用MD5 

            byte[] Change = s1.ComputeHash(Original);//進行加密 

            return Convert.ToBase64String(Change);//將加密後的字串從byte[]轉回string

        }

   

    }
}