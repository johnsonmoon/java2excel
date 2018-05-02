package com.github.johnsonmoon.java2excel.core.entity.formatted.model.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Attribute annotation for entity field which is going to generate excel
 * <p>
 * Created by xuyh at 2018/1/3 13:33.
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ ElementType.FIELD })
public @interface Attribute {
	/**
	 * {@link com.github.johnsonmoon.java2excel.core.entity.formatted.model.Attribute#attrName}
	 *
	 * @return
	 */
	String attrName();

	/**
	 * {@link com.github.johnsonmoon.java2excel.core.entity.formatted.model.Attribute#attrType}
	 *
	 * @return
	 */
	String attrType() default "";

	/**
	 * {@link com.github.johnsonmoon.java2excel.core.entity.formatted.model.Attribute#formatInfo}
	 *
	 * @return
	 */
	String formatInfo() default "";

	/**
	 * {@link com.github.johnsonmoon.java2excel.core.entity.formatted.model.Attribute#defaultValue}
	 *
	 * @return
	 */
	String defaultValue() default "";

	/**
	 * {@link com.github.johnsonmoon.java2excel.core.entity.formatted.model.Attribute#unit}
	 *
	 * @return
	 */
	String unit() default "";
}
