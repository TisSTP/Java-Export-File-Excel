package util;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Utility {
  /**
   * Checks if is valid.
   *
   * @param obj the obj
   * @return true, if is valid
   */
  public static boolean isValid(Object obj) {
    if (obj != null) {
      if (obj instanceof ArrayList) {
        ArrayList<?> list = (ArrayList<?>) obj;
        if (list.size() == 0) {
          return false;
        }
      } else if (obj instanceof Object[]) {
        Object[] list = (Object[]) obj;
        if (list.length == 0) {
          return false;
        }
      } else if (obj instanceof List) {
        List<?> list = (List<?>) obj;
        if (list.size() == 0) {
          return false;
        }
      } else if (obj instanceof Collection) {
        Collection<?> collection = (Collection<?>) obj;
        if (collection.size() == 0) {
          return false;
        }
      } else if (obj instanceof BigDecimal) {
        BigDecimal bigDecimalData = (BigDecimal) obj;
        if (bigDecimalData == null || bigDecimalData.compareTo(BigDecimal.ZERO) <= 0) {
          return false;
        }
      } else if (obj instanceof Long) {
        Long longData = (Long) obj;
        if (longData == null || longData <= 0) {
          return false;
        }
      } else if (obj instanceof Double) {
        Double doubleData = (Double) obj;
        if (doubleData == null || doubleData <= 0) {
          return false;
        }
      } else if (obj instanceof Integer) {
        Integer intData = (Integer) obj;
        if (intData == null || intData <= 0) {
          return false;
        }
      } else if (obj instanceof String) {
        String intData = (String) obj;
        if (intData == null || intData.trim().length() == 0 || intData.equalsIgnoreCase("null")) {
          return false;
        }
      } else if (obj instanceof StringBuffer) {
        StringBuffer intData = (StringBuffer) obj;
        if (intData == null || intData.toString().trim().length() == 0
            || intData.toString().equalsIgnoreCase("null")) {
          return false;
        }
      } else if (obj instanceof Map) {
        Map<?, ?> intData = (Map<?, ?>) obj;
        if (intData == null || intData.size() == 0) {
          return false;
        }
      } else if (obj instanceof HashMap) {
        HashMap<?, ?> intData = (HashMap<?, ?>) obj;
        if (intData == null || intData.size() == 0) {
          return false;
        }
      }
      return true;
    }
    return false;
  }
}
