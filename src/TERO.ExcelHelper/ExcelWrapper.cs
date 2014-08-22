using System;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace TERO.ExcelHelper
{
  public class ExcelWrapper : TERO.BaseManager.BaseDataManager
  {
    #region Fields
    
    public IWorkbook ExcelBook;    // Declaring variable in type of Excel Workbook
    public ISheet ExcelSheet;      // Declaring variable in type of Excel WorkSheet
    
    #endregion

    #region Constructor

    public ExcelWrapper()
    {
      OpenedExcelFilename = string.Empty;
      LastGetValueIsNull = true;
    }

    #endregion

    #region Properties

    public string LastCellRange { get; set; }

    public bool LastGetValueIsNull { get; private set; }

    public string OpenedExcelFilename { get; private set; }

    public string ExcelFirstColumn
    {
      get { return "A"; }
    }

    public string ExcelLastColumn
    {
      get { return "IV"; }
    }

    public int ExcelFirstColumnNo
    {
      get { return 1; }
    }

    public int ExcelLastColumnNo
    {
      get { return 256; }
    }

    public int ExcelFirstRow
    {
      get { return 1; }
    }

    public int ExcelLastRow
    {
      get { return 65536; }
    }

    #endregion

    #region Enums

    public enum WorksheetPosition
    {
      Default = 1,
      First = 2,
      Last = 3
    };

    public enum HorizontalAlignment
    {
      General = 1,
      Left = 2,
      Center = 3,
      Right = 4,
      Fill = 5,
      Justify = 6,
      CenterAcrossSelection = 7,
      Distributed = 8
    };

    public enum VerticalAlignment
    {
      Top = 1,
      Center = 2,
      Bottom = 3,
      Justify = 4,
      Distributed = 5
    };

    #endregion

    #region Methods

    public string GetCellValue(int row, int col)
    {
      return GetCellValue(row, col, ExcelSheet);
    }

    public string GetCellValue(int row, int col, ISheet sheet)
    {
      string value = null;

      try
      {
        var r = sheet.GetRow(row);
        if (r != null)
        {
          var cell = r.GetCell(col, MissingCellPolicy.CREATE_NULL_AS_BLANK);
          value = GetCellValue(cell);
        }
      }
      catch (Exception ex)
      {
        OnStatus(ex.Message);
      }

      return value;
    }

    public string GetCellValue(ICell cell)
    {
      return GetCellValue(cell, null);
    }

    public string GetCellValue(ICell cell, CellType? cellType)
    {
      return GetCellValueType(cell, cellType).Value;
    }
    



    public CellValueType.CellType GetCellType(int row, int col)
    {
      return GetCellType(row, col, ExcelSheet);
    }

    public CellValueType.CellType GetCellType(int row, int col, ISheet sheet)
    {
      var value = CellValueType.CellType.Unknown;

      try
      {
        var r = sheet.GetRow(row);
        if (r != null)
        {
          var cell = r.GetCell(col, MissingCellPolicy.CREATE_NULL_AS_BLANK);
          value = GetCellType(cell);
        }
      }
      catch (Exception ex)
      {
        OnStatus(ex.Message);
      }

      return value;
    }

    public CellValueType.CellType GetCellType(ICell cell)
    {
      return GetCellType(cell, null);
    }

    public CellValueType.CellType GetCellType(ICell cell, CellType? cellType)
    {
      return GetCellValueType(cell, cellType).ValueCellType;
    }
    

    public CellValueType GetCellValueType(ICell cell)
    {
      return GetCellValueType(cell, null);
    }

    public CellValueType GetCellValueType(ICell cell, CellType? cellType)
    {
      var cvt = new CellValueType() {Value = ""};
      if (cellType == null) { cellType = cell.CellType; }

      try
      {
        switch (cellType)
        {
          case CellType.Boolean:
            cvt.Value = cell.BooleanCellValue.ToString();
            cvt.ValueCellType = CellValueType.CellType.Boolean;
            break;
          case CellType.Numeric:
            if (DateUtil.IsCellDateFormatted(cell))
            {
              cvt.Value = DateUtil.GetJavaDate(cell.NumericCellValue).ToString();
              cvt.ValueCellType = CellValueType.CellType.Date;
            }
            else
            {
              cvt.Value = cell.NumericCellValue.ToString();
              cvt.ValueCellType = CellValueType.CellType.Numeric;
            }
            break;
          case CellType.Formula:
            var formulaCvt = GetCellValueType(cell, cell.CachedFormulaResultType);
            cvt.Value = formulaCvt.Value;
            switch (formulaCvt.ValueCellType)
            {
              case CellValueType.CellType.Boolean:
                cvt.ValueCellType = CellValueType.CellType.Formula_Boolean;
                break;
              case CellValueType.CellType.Date:
                cvt.ValueCellType = CellValueType.CellType.Formula_Date;
                break;
              case CellValueType.CellType.Error:
                cvt.ValueCellType = CellValueType.CellType.Formula_Error;
                break;
              case CellValueType.CellType.Numeric:
                cvt.ValueCellType = CellValueType.CellType.Formula_Numeric;
                break;
              case CellValueType.CellType.String:
                cvt.ValueCellType = CellValueType.CellType.Formula_String;
                break;
              default:
                break;
            }
            break;
          case CellType.Error:
            cvt.Value = "#ERROR!";
            cvt.ValueCellType = CellValueType.CellType.Error;
            break;
          default:
            cvt.Value = cell.StringCellValue;
            cvt.ValueCellType = CellValueType.CellType.String;
            break;
        }
      }
      catch (Exception ex)
      {
        OnStatus(ex.Message);
      }

      return cvt;
    }




    public void Open(string filename)
    {
      if (!File.Exists(filename))
      {
        throw new FileNotFoundException("Unable to locate file.", filename);
      }

      using (var fs = new FileStream(filename, FileMode.Open))
      {
        ExcelBook = WorkbookFactory.Create(fs);
        ExcelSheet = ExcelBook.GetSheetAt(ExcelBook.ActiveSheetIndex);
        fs.Close();
      }

      this.OpenedExcelFilename = filename;
    }

    public void Create(string filename)
    {
      if (File.Exists(filename))
      {
        throw new ApplicationException("File already exists!");
      }

      ExcelBook = new HSSFWorkbook();
      ExcelBook.CreateSheet("Sheet1");
      this.OpenedExcelFilename = filename;
    }

    public void Close()
    {
      ExcelSheet = null;
      ExcelBook = null;
      this.OpenedExcelFilename = "";
    }

    public void Save()
    {
      using (var fs = new FileStream(this.OpenedExcelFilename, FileMode.Create))
      {
        ExcelBook.Write(fs);
        fs.Close();
      }
    }

    public void SaveAs(string filename)
    {
      using (var fs = new FileStream(filename, FileMode.Create))
      {
        ExcelBook.Write(fs);
        fs.Close();
      }
      Open(filename);
    }


    public void Test()
    {
      var filename = Path.Combine(Environment.CurrentDirectory, "test.xlsx");
      ExcelBook = WorkbookFactory.Create(filename);
      ExcelSheet = ExcelBook.GetSheetAt(0);

      //0
      //3 to 6
      var firstRow = 1;
      var lastRow = 2;
      var firsCol = CellReference.ConvertColStringToIndex("E");
      var lastCol = CellReference.ConvertColStringToIndex("S");

      for (int row = firstRow; row <= lastRow; row++)
      {
        for (var col = firsCol; col <= lastCol; col++)
        {
          var value = GetCellValue(row, col, ExcelSheet);
          var valueType = GetCellType(row, col, ExcelSheet);
          OnStatus("[{0},{1}], {2}, '{3}'", row, col, value, valueType);
        }
        OnStatus("");
      }
    }

    #endregion
  }
}