package com.dao.chu.excel;

import java.math.BigDecimal;

/** 
 * 
 * ѧ��ͳ����Ϣ 
 * @author daochuwenziyao 
 * @see [�����/����] 
 * @since [��Ʒ/ģ��汾] 
 */
public class StudentStatistics 
{ 
 private Integer id; 
 @IsNeeded
 private BigDecimal totalGrade; 
 @IsNeeded
 private BigDecimal avgGrade; 
   
 @Override
 public String toString() 
 { 
  return "StudentStatistics [id=" + id + ", totalGrade=" + totalGrade + ", avgGrade=" + avgGrade + "]"; 
 } 
 public Integer getId() 
 { 
  return id; 
 } 
 public void setId(Integer id) 
 { 
  this.id = id; 
 } 
 public BigDecimal getTotalGrade() 
 { 
  return totalGrade; 
 } 
 public void setTotalGrade(BigDecimal totalGrade) 
 { 
  this.totalGrade = totalGrade; 
 } 
 public BigDecimal getAvgGrade() 
 { 
  return avgGrade; 
 } 
 public void setAvgGrade(BigDecimal avgGrade) 
 { 
  this.avgGrade = avgGrade; 
 } 
   
}
