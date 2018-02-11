package xuyihao.java2excel.core.entity.custom.map.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * annotation for [column:fieldName] mapping which is going to generate/read excel
 * <p>
 * Created by xuyh at 2018/2/10 10:42.
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ ElementType.FIELD })
public @interface Column {
	/**
	 * Set column for the field.
	 */
	int column();
}
