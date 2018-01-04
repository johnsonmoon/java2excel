package xuyihao.java2excel.core.entity.annotation;

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
	 * {@link xuyihao.java2excel.core.entity.model.Attribute#attrName}
	 *
	 * @return
	 */
	String attrName();

	/**
	 * {@link xuyihao.java2excel.core.entity.model.Attribute#attrType}
	 *
	 * @return
	 */
	String attrType();

	/**
	 * {@link xuyihao.java2excel.core.entity.model.Attribute#formatInfo}
	 *
	 * @return
	 */
	String formatInfo() default "";

	/**
	 * {@link xuyihao.java2excel.core.entity.model.Attribute#defaultValue}
	 *
	 * @return
	 */
	String defaultValue() default "";

	/**
	 * {@link xuyihao.java2excel.core.entity.model.Attribute#unit}
	 *
	 * @return
	 */
	String unit() default "";
}
