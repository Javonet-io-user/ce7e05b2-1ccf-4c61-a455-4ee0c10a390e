package cubaRG;

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
import cubaRG.*;
import Microsoft.Office.Interop.Excel.*;
import jio.System.IO.*;
import jio.System.Drawing.*;

public class ClassExcel {
  protected NObject javonetHandle;
  /** GetFiled */
  public Application getm_objExcel() {
    try {
      Object res = javonetHandle.<NObject>get("m_objExcel");
      if (res == null) return null;
      return ConvertToConcreteInterfaceImplementation(res);
    } catch (JavonetException _javonetException) {
      _javonetException.printStackTrace();
      return null;
    }
  }
  /** SetFiled */

  public void setm_objExcel(Application param) {
    try {
      javonetHandle.set("m_objExcel", param);
    } catch (JavonetException _javonetException) {
      _javonetException.printStackTrace();
    }
  }
  /** GetFiled */

  public Workbooks getm_objBooks() {
    try {
      Object res = javonetHandle.<NObject>get("m_objBooks");
      if (res == null) return null;
      return ConvertToConcreteInterfaceImplementation(res);
    } catch (JavonetException _javonetException) {
      _javonetException.printStackTrace();
      return null;
    }
  }
  /** SetFiled */

  public void setm_objBooks(Workbooks param) {
    try {
      javonetHandle.set("m_objBooks", param);
    } catch (JavonetException _javonetException) {
      _javonetException.printStackTrace();
    }
  }
  /** GetFiled */

  public Workbook getm_objBook() {
    try {
      Object res = javonetHandle.<NObject>get("m_objBook");
      if (res == null) return null;
      return ConvertToConcreteInterfaceImplementation(res);
    } catch (JavonetException _javonetException) {
      _javonetException.printStackTrace();
      return null;
    }
  }
  /** SetFiled */

  public void setm_objBook(Workbook param) {
    try {
      javonetHandle.set("m_objBook", param);
    } catch (JavonetException _javonetException) {
      _javonetException.printStackTrace();
    }
  }
  /** GetFiled */

  public Sheets getm_objSheets() {
    try {
      Object res = javonetHandle.<NObject>get("m_objSheets");
      if (res == null) return null;
      return ConvertToConcreteInterfaceImplementation(res);
    } catch (JavonetException _javonetException) {
      _javonetException.printStackTrace();
      return null;
    }
  }
  /** SetFiled */

  public void setm_objSheets(Sheets param) {
    try {
      javonetHandle.set("m_objSheets", param);
    } catch (JavonetException _javonetException) {
      _javonetException.printStackTrace();
    }
  }
  /** GetFiled */

  public Color getColorYelloGreen() {
    try {
      Object res = javonetHandle.<NObject>get("ColorYelloGreen");
      if (res == null) return null;
      return new Color((NObject) res);
    } catch (JavonetException _javonetException) {
      _javonetException.printStackTrace();
      return null;
    }
  }
  /** SetFiled */

  public void setColorYelloGreen(Color param) {
    try {
      javonetHandle.set("ColorYelloGreen", param);
    } catch (JavonetException _javonetException) {
      _javonetException.printStackTrace();
    }
  }
  /** GetFiled */

  public Color getColorDelicateOrange() {
    try {
      Object res = javonetHandle.<NObject>get("ColorDelicateOrange");
      if (res == null) return null;
      return new Color((NObject) res);
    } catch (JavonetException _javonetException) {
      _javonetException.printStackTrace();
      return null;
    }
  }
  /** SetFiled */

  public void setColorDelicateOrange(Color param) {
    try {
      javonetHandle.set("ColorDelicateOrange", param);
    } catch (JavonetException _javonetException) {
      _javonetException.printStackTrace();
    }
  }
  /** GetFiled */

  public Color getColorDelicateBlue() {
    try {
      Object res = javonetHandle.<NObject>get("ColorDelicateBlue");
      if (res == null) return null;
      return new Color((NObject) res);
    } catch (JavonetException _javonetException) {
      _javonetException.printStackTrace();
      return null;
    }
  }
  /** SetFiled */

  public void setColorDelicateBlue(Color param) {
    try {
      javonetHandle.set("ColorDelicateBlue", param);
    } catch (JavonetException _javonetException) {
      _javonetException.printStackTrace();
    }
  }
  /** GetFiled */

  public Color getColorDelicateGreen() {
    try {
      Object res = javonetHandle.<NObject>get("ColorDelicateGreen");
      if (res == null) return null;
      return new Color((NObject) res);
    } catch (JavonetException _javonetException) {
      _javonetException.printStackTrace();
      return null;
    }
  }
  /** SetFiled */

  public void setColorDelicateGreen(Color param) {
    try {
      javonetHandle.set("ColorDelicateGreen", param);
    } catch (JavonetException _javonetException) {
      _javonetException.printStackTrace();
    }
  }
  /** GetFiled */

  public Color getColorFooter() {
    try {
      Object res = javonetHandle.<NObject>get("ColorFooter");
      if (res == null) return null;
      return new Color((NObject) res);
    } catch (JavonetException _javonetException) {
      _javonetException.printStackTrace();
      return null;
    }
  }
  /** SetFiled */

  public void setColorFooter(Color param) {
    try {
      javonetHandle.set("ColorFooter", param);
    } catch (JavonetException _javonetException) {
      _javonetException.printStackTrace();
    }
  }
  /** GetFiled */

  public Color getColorDarkBlue() {
    try {
      Object res = javonetHandle.<NObject>get("ColorDarkBlue");
      if (res == null) return null;
      return new Color((NObject) res);
    } catch (JavonetException _javonetException) {
      _javonetException.printStackTrace();
      return null;
    }
  }
  /** SetFiled */

  public void setColorDarkBlue(Color param) {
    try {
      javonetHandle.set("ColorDarkBlue", param);
    } catch (JavonetException _javonetException) {
      _javonetException.printStackTrace();
    }
  }
  /** GetFiled */

  public Color getColorTab() {
    try {
      Object res = javonetHandle.<NObject>get("ColorTab");
      if (res == null) return null;
      return new Color((NObject) res);
    } catch (JavonetException _javonetException) {
      _javonetException.printStackTrace();
      return null;
    }
  }
  /** SetFiled */

  public void setColorTab(Color param) {
    try {
      javonetHandle.set("ColorTab", param);
    } catch (JavonetException _javonetException) {
      _javonetException.printStackTrace();
    }
  }

  public ClassExcel() {
    try {
      javonetHandle = Javonet.New("cubaRG.ClassExcel");
    } catch (JavonetException _javonetException) {
      _javonetException.printStackTrace();
    }
  }

  public ClassExcel(NObject handle) {
    this.javonetHandle = handle;
  }

  public void setJavonetHandle(NObject handle) {
    this.javonetHandle = handle;
  }
  /** Method */

  public void MakeCompliant(Worksheet _m_objSheet_Data) {
    try {
      javonetHandle.invoke("MakeCompliant", _m_objSheet_Data);
    } catch (JavonetException _javonetException) {
      _javonetException.printStackTrace();
    }
  }
  /** Method */

  public void MakeReport(
      Worksheet _m_objSheet_Data, Worksheet _m_objSheet_Report, java.lang.Boolean flg) {
    try {
      javonetHandle.invoke("MakeReport", _m_objSheet_Data, _m_objSheet_Report, flg);
    } catch (JavonetException _javonetException) {
      _javonetException.printStackTrace();
    }
  }
  /** Method */

  public void CreateNewXLSFile(
      java.lang.Boolean isVisible, java.lang.String inputFile, java.lang.String outputFile) {
    try {
      javonetHandle.invoke("CreateNewXLSFile", isVisible, inputFile, outputFile);
    } catch (JavonetException _javonetException) {
      _javonetException.printStackTrace();
    }
  }
  /** Method */

  public void CreateWorksheetAuto() {
    try {
      javonetHandle.invoke("CreateWorksheetAuto");
    } catch (JavonetException _javonetException) {
      _javonetException.printStackTrace();
    }
  }
  /** Method */

  public void CreateWorksheetManual() {
    try {
      javonetHandle.invoke("CreateWorksheetManual");
    } catch (JavonetException _javonetException) {
      _javonetException.printStackTrace();
    }
  }
  /** Method */

  public void SaveWorkBook(FileInfo fi) {
    try {
      javonetHandle.invoke("SaveWorkBook", fi);
    } catch (JavonetException _javonetException) {
      _javonetException.printStackTrace();
    }
  }
  /** Method */

  public void closeApplication() {
    try {
      javonetHandle.invoke("closeApplication");
    } catch (JavonetException _javonetException) {
      _javonetException.printStackTrace();
    }
  }

  static {
    try {
      Activation.initializeJavonet();
    } catch (java.lang.Exception e) {
      e.printStackTrace();
    }
  }
}
