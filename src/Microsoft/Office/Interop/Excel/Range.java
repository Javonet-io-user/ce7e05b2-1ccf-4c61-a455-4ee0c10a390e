package Microsoft.Office.Interop.Excel;

import Common.Activation;
import static Common.Helper.Convert;
import static Common.Helper.getGetObjectName;
import static Common.Helper.getReturnObjectName;
import static Common.Helper.ConvertToConcreteInterfaceImplementation;
import Common.Helper;
import com.javonet.Javonet;
import com.javonet.JavonetException;
import com.javonet.JavonetFramework;
import com.javonet.api.NObject;
import com.javonet.api.NEnum;
import com.javonet.api.keywords.NRef;
import com.javonet.api.keywords.NOut;
import com.javonet.api.NControlContainer;
import java.util.concurrent.atomic.AtomicReference;
import java.util.Iterator;
import java.lang.*;
import Microsoft.Office.Interop.Excel.*;
import jio.System.*;
import jio.System.Collections.*;

public interface Range extends IEnumerable {
  public Object AdvancedFilter(
      XlFilterAction Action, Object CriteriaRange, Object CopyToRange, Object Unique);

  public Object AutoFilter(
      Object Field,
      Object Criteria1,
      XlAutoFilterOperator Operator,
      Object Criteria2,
      Object VisibleDropDown);

  public Object AutoFit();

  public Object Copy(Object Destination);

  public Object Cut(Object Destination);

  public Object Delete(Object Shift);

  public Object Insert(Object Shift, Object CopyOrigin);

  public IEnumerator GetEnumerator();

  public Object Select();

  public Object Sort(
      Object Key1,
      XlSortOrder Order1,
      Object Key2,
      Object Type,
      XlSortOrder Order2,
      Object Key3,
      XlSortOrder Order3,
      XlYesNoGuess Header,
      Object OrderCustom,
      Object MatchCase,
      XlSortOrientation Orientation,
      XlSortMethod SortMethod,
      XlSortDataOption DataOption1,
      XlSortDataOption DataOption2,
      XlSortDataOption DataOption3);

  public Range SpecialCells(XlCellType Type, Object Value);

  public Object PasteSpecial(
      XlPasteType Paste, XlPasteSpecialOperation Operation, Object SkipBlanks, Object Transpose);
}
