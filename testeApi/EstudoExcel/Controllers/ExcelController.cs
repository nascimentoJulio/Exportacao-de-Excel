using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using EstudoExcel.models;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Style;



namespace EstudoExcel.Controllers
{
  [Route("download")]
  public class ExcelController : ControllerBase
  {
    private int rowIndex;
    private int colIndex;

    private ExcelWorksheet worksheet;

    private ExcelPackage package;

    [HttpGet]
    public FileResult Dowload()
    {

      CreateFile("test.xlsx");

      var lst = new List<object>(){
        new ExcelModel{Id = 3, Value = "dsddsd", date = DateTime.Now},
        new ExcelModel{Id = 3, Value = "dsddsd", date = DateTime.Now},
        new ExcelModel{Id = 3, Value = "dsddsd", date = DateTime.Now}
      };
      InsertSheet(lst, "calc");
      var file = GetFile();
      return File(file, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "meuarquivo.xlsx");
    }

    //Recebe um nome para o arquivo e cria o arquivo
    private void CreateFile(string fileName)
    {
      if (fileName.Equals(string.Empty))
      {
        package = new ExcelPackage();
      }
      else
      {
        FileInfo fileInfo = new FileInfo(fileName);
        if (!fileInfo.Extension.Equals(".xlsx"))
        {
          fileInfo = null;
          throw new Exception(string.Format("Tipo de arquivo inválido, {0} esperado.", "xlsx"));
        }

        if (fileInfo.Exists)
        {
          fileInfo.Delete();
        }

        package = new ExcelPackage(fileInfo);
      }
    }

    //Obtém o arquivo recém criado
    private byte[] GetFile()
    {
      Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
      byte[] result = package.GetAsByteArray();
      package.Dispose();
      return result;
    }

    // cria a folha
    private void InsertSheet(List<object> lstData, string sheetName = "")
    {
      if (lstData == null || lstData.Count.Equals(0))
      {
        worksheet = package.Workbook.Worksheets.Add("Grid");
        return;
      }

      worksheet = package.Workbook.Worksheets.Add(sheetName.Equals(string.Empty) ? lstData[0].GetType().Name : sheetName);
      FillHeader(lstData[0]);
      lstData.ForEach(d => FillValues(d));
    }

    private void FillHeader(object data)
    {
      rowIndex = 1;
      FillValues(data);
    }

    //manda os valores para as celulas
    private void FillValues(object data)
    {
      colIndex = 1;
      foreach (PropertyInfo property in data.GetType().GetProperties())
      {
        SetCellProperties(property, data);
      }
      rowIndex++;
    }

    private void FormatDetailCell(ExcelRange excelRange, string dataType)
    {
      string numberFormat = string.Empty;
      switch (dataType)
      {
        case "Int32":
        case "Int64":
          numberFormat = "#,##0";
          break;
        case "Double":
          numberFormat = "#,##0.00";
          break;
        case "DateTime":
          numberFormat = "dd/MM/yyyy hh:mm";
          break;
        default:
          break;
      }
      excelRange.Style.Numberformat.Format = numberFormat;
    }

    private void FormatHeaderCell(ExcelRange excelRange)
    {
      excelRange.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
      excelRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
      excelRange.Style.Fill.BackgroundColor.SetColor(Color.White);
      excelRange.Style.Font.Color.SetColor(Color.White);
      excelRange.Style.Font.Bold = true;
      excelRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
      excelRange.Style.Border.Right.Color.SetColor(Color.White);
      excelRange.AutoFitColumns();
    }

    // define as propiedades da celula
    private void SetCellProperties(PropertyInfo property, object data)
    {
      ExcelRange excelRange = worksheet.Cells[rowIndex, colIndex];

      ExportAttribute exportAttribute = property.GetCustomAttribute<ExportAttribute>();
      if (exportAttribute != null && exportAttribute.DiscardColumn)
      {
        return;
      }

      if (rowIndex.Equals(1))
      {
        excelRange.Value = exportAttribute != null ? exportAttribute.ColumnName : property.Name;
        FormatHeaderCell(excelRange);
      }
      else
      {
        excelRange.Value = property.GetValue(data, null);
        FormatDetailCell(excelRange, property.PropertyType.Name);
      }

      colIndex++;
    }

  
  }
}