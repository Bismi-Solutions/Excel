/*
 * Copyright (c) 2019. Bismi Solutions
 *
 * https://bismi.solutions/
 * support@bismi.solutions
 * sulfikar.ali.nazar@gmail.com
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */

package solutions.bismi.excel;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * @author Sulfikar Ali Nazar
 */
public abstract class Common {
    private static Logger log = LogManager.getLogger(Common.class);

    private static IndexedColors[] values=null;
    private static Map<String,Short> map = new HashMap<String,Short>();

    public static void updateIndexedValues(){
        if(values==null){
            values = getEnumValues(IndexedColors.class);
            for (IndexedColors ele:values) {
                map.put(ele.toString(),ele.getIndex());
            }
        }

    }
    public static short  getColorCode(String color){
        updateIndexedValues();

        String _colors=Arrays.toString(values);
        if(map.containsKey(color.toUpperCase().trim())){
            return map.get(color.toUpperCase().trim());
        }else{
            log.info("Supported string constant colors are :" + _colors);
            log.info("Hex code can also be passed if you need more colors...");
            return 0;
        }



    }

    private static <E extends Enum> E[] getEnumValues(Class<E> enumClass) {

        try{
            Field fld = enumClass.getDeclaredField("$VALUES");
            fld.setAccessible(true);
            Object o = fld.get(null);
            return (E[]) o;
        }catch(Exception e){
            log.info("Error in getting color enumerators");
        return null;
        }

    }

}
