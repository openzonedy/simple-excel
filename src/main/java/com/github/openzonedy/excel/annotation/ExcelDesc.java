package com.github.openzonedy.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelDesc {

    String value() default "";

    /**
     * 读取Excel数据到对象属性指定设值方法
     * 在设值的过程中valueSetter优先级最高
     *
     * @return
     */
    String valueSetter() default "";

    /**
     * 写出对象到Excel指定取值的方法
     * 在取值的过程中valueGetter的优先级最高
     *
     * @return
     */
    String valueGetter() default "";

    /**
     * 时间格式化和反格式化规则
     *
     * @return
     */
    String timePattern() default "";

    /**
     * 取枚举中的某个属性值(如果指定了，默认读取和写出都用这个属性，可以用valueSetter/valueGetter覆盖)
     *
     * @return
     */
    String enumValue() default "";

    /**
     * 布尔值读写时映射关系
     *
     * @return
     */
    String booleanPattern() default "TRUE/FALSE";

    /**
     * 排序
     *
     * @return
     */
    int order() default 1;
}
