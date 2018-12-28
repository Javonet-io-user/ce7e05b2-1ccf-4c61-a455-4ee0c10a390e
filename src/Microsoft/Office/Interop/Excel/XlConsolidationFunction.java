package Microsoft.Office.Interop.Excel;

public enum XlConsolidationFunction {
  xlDistinctCount(11L),
  xlUnknown(1000L),
  xlVarP(-4165L),
  xlVar(-4164L),
  xlSum(-4157L),
  xlStDevP(-4156L),
  xlStDev(-4155L),
  xlProduct(-4149L),
  xlMin(-4139L),
  xlMax(-4136L),
  xlCountNums(-4113L),
  xlCount(-4112L),
  xlAverage(-4106L),
  ;
  private long numVal;

  XlConsolidationFunction(long numVal) {
    this.numVal = numVal;
  }

  public long getNumVal() {
    return numVal;
  }
}
