package annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Target;

@ExcelAnnotation
@Target(ElementType.FIELD)
public @interface ExportExcel {
    /*
    导出的文件的名字
     */
    String fileName() default "";

    /*
    导出文件的方式
     */
    String exportWay() default "";


}
