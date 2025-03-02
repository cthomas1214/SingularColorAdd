using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;


class ColorAdd
{
    static void Main()
    {
        //Folder path and connection string
        string folderPath = @"E:\ExcelPracticeFiles\excelfiles";
        string connectionString = "";
        ExcelProcessor processor = new ExcelProcessor(folderPath, connectionString);
        processor.ProcessFiles();
    }
}