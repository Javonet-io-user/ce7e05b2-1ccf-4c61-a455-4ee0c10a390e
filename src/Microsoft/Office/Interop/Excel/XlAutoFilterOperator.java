package Microsoft.Office.Interop.Excel;

public enum XlAutoFilterOperator {
  xlAnd(1L),
  xlOr(2L),
  xlTop10Items(3L),
  xlBottom10Items(4L),
  xlTop10Percent(5L),
  xlBottom10Percent(6L),
  xlFilterValues(7L),
  xlFilterCellColor(8L),
  xlFilterFontColor(9L),
  xlFilterIcon(10L),
  xlFilterDynamic(11L),
  xlFilterNoFill(12L),
  xlFilterAutomaticFontColor(13L),
  xlFilterNoIcon(14L),
  ;
  private long numVal;

  XlAutoFilterOperator(long numVal) {
    this.numVal = numVal;
  }

  public long getNumVal() {
    return numVal;
  }
}
