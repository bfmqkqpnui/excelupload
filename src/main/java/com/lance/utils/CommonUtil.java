package com.lance.utils;

import java.util.List;
import java.util.Map;

/**
 * @author
 * @create 2018-02-04 9:20
 **/
public class CommonUtil {
    public static boolean isExist(Object obj) {
        boolean flag = false;
        if (null != obj) {
            if (obj instanceof List) {
                List list = (List) obj;
                if (null != list && list.size() > 0) {
                    flag = true;
                } else if (obj instanceof Map) {
                    Map map = (Map) obj;
                    if (null != map && map.size() > 0) {
                        flag = true;
                    }
                }
            }
        }
        return flag;
    }
}
