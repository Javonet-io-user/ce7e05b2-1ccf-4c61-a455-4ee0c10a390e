package Microsoft.Office.Interop.Excel;

public enum XlPasteType {
  xlPasteValidation(6L),
  xlPasteAllExceptBorders(7L),
  xlPasteColumnWidths(8L),
  xlPasteFormulasAndNumberFormats(11L),
  xlPasteValuesAndNumberFormats(12L),
  xlPasteAllUsingSourceTheme(13L),
  xlPasteAllMergingConditionalFormats(14L),
  xlPasteValues(-4163L),
  xlPasteComments(-4144L),
  xlPasteFormulas(-4123L),
  xlPasteFormats(-4122L),
  xlPasteAll(-4104L),
  ;
  private long numVal;

  XlPasteType(long numVal) {
    this.numVal = numVal;
  }

  public long getNumVal() {
    return numVal;
  }
}
