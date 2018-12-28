package Microsoft.Office.Interop.Excel;

public enum XlCellType {
  xlCellTypeConstants(2L),
  xlCellTypeBlanks(4L),
  xlCellTypeLastCell(11L),
  xlCellTypeVisible(12L),
  xlCellTypeSameValidation(-4175L),
  xlCellTypeAllValidation(-4174L),
  xlCellTypeSameFormatConditions(-4173L),
  xlCellTypeAllFormatConditions(-4172L),
  xlCellTypeComments(-4144L),
  xlCellTypeFormulas(-4123L),
  ;
  private long numVal;

  XlCellType(long numVal) {
    this.numVal = numVal;
  }

  public long getNumVal() {
    return numVal;
  }
}
