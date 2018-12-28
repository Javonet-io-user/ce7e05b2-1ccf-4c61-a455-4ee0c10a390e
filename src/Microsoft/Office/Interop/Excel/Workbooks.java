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

public interface Workbooks extends IEnumerable {
  public Workbook Open(
      java.lang.String Filename,
      Object UpdateLinks,
      Object ReadOnly,
      Object Format,
      Object Password,
      Object WriteResPassword,
      Object IgnoreReadOnlyRecommended,
      Object Origin,
      Object Delimiter,
      Object Editable,
      Object Notify,
      Object Converter,
      Object AddToMru,
      Object Local,
      Object CorruptLoad);
}
