package com.github.johnsonmoon.java2excel.core.entity.formatted.model.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Model annotation for entity which is going to generate excel
 * <p>
 * Created by xuyh at 2018/1/3 11:46.
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ ElementType.TYPE })
public @interface Model {
	/**
	 * {@link xuyihao.java2excel.core.entity.formatted.model.Model#name}
	 *
	 * @return
	 */
	String name();
}
