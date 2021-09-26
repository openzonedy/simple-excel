package com.github.openzonedy.excel;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelDesc {

    String valueSetter() default "";

    String valueGetter() default "";

    String timePattern() default "";

    String booleanPattern() default "TRUE/FALSE";
}
