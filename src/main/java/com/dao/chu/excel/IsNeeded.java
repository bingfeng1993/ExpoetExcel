package com.dao.chu.excel;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 
 * �Ƿ���Ҫ�ӽ���excel��ֵ
 * 
 * @author daochuwenziyao
 * @see [�����/����]
 * @since [��Ʒ/ģ��汾]
 */
@Retention(value = RetentionPolicy.RUNTIME)
@Target(value = { ElementType.FIELD })
public @interface IsNeeded {

	/**
	 * �Ƿ���Ҫ�ӽ���excel��ֵ
	 * 
	 * @return true:��Ҫ false:����Ҫ
	 * @see [�ࡢ��#��������#��Ա]
	 */
	boolean isNeeded() default true;
}
