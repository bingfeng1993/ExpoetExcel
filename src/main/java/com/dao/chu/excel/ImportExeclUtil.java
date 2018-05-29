package com.dao.chu.excel; 
  
import java.io.IOException; 
import java.io.InputStream; 
import java.lang.reflect.Field; 
import java.math.BigDecimal; 
import java.text.DecimalFormat; 
import java.text.ParseException; 
import java.text.SimpleDateFormat; 
import java.util.ArrayList; 
import java.util.Date; 
import java.util.List; 
import java.util.Locale; 
  
import org.apache.commons.beanutils.PropertyUtils; 
import org.apache.commons.lang.StringUtils; 
import org.apache.poi.hssf.usermodel.HSSFCell; 
import org.apache.poi.hssf.usermodel.HSSFWorkbook; 
import org.apache.poi.ss.usermodel.Cell; 
import org.apache.poi.ss.usermodel.Row; 
import org.apache.poi.ss.usermodel.Sheet; 
import org.apache.poi.ss.usermodel.Workbook; 
import org.apache.poi.xssf.usermodel.XSSFWorkbook; 
  
/** 
 * 
 * excel��ȡ������ 
 * 
 * @author daochuwenziyao 
 * @see [�����/����] 
 * @since [��Ʒ/ģ��汾] 
 */
public class ImportExeclUtil 
{ 
   
 private static int totalRows = 0;// ������ 
   
 private static int totalCells = 0;// ������ 
   
 private static String errorInfo;// ������Ϣ 
   
 /** �޲ι��췽�� */
 public ImportExeclUtil() 
 { 
 } 
   
 public static int getTotalRows() 
 { 
  return totalRows; 
 } 
   
 public static int getTotalCells() 
 { 
  return totalCells; 
 } 
   
 public static String getErrorInfo() 
 { 
  return errorInfo; 
 } 
   
 /** 
  * 
  * ��������ȡExcel�ļ� 
  * 
  * 
  * @param inputStream 
  * @param isExcel2003 
  * @return 
  * @see [�ࡢ��#��������#��Ա] 
  */
 public List<List<String>> read(InputStream inputStream, boolean isExcel2003) 
  throws IOException 
 { 
    
  List<List<String>> dataLst = null; 
    
  /** ���ݰ汾ѡ�񴴽�Workbook�ķ�ʽ */
  Workbook wb = null; 
    
  if (isExcel2003) 
  { 
   wb = new HSSFWorkbook(inputStream); 
  } 
  else
  { 
   wb = new XSSFWorkbook(inputStream); 
  } 
  dataLst = readDate(wb); 
    
  return dataLst; 
 } 
   
 /** 
  * 
  * ��ȡ���� 
  * 
  * @param wb 
  * @return 
  * @see [�ࡢ��#��������#��Ա] 
  */
 private List<List<String>> readDate(Workbook wb) 
 { 
    
  List<List<String>> dataLst = new ArrayList<List<String>>(); 
    
  /** �õ���һ��shell */
  Sheet sheet = wb.getSheetAt(0); 
    
  /** �õ�Excel������ */
  totalRows = sheet.getPhysicalNumberOfRows(); 
    
  /** �õ�Excel������ */
  if (totalRows >= 1 && sheet.getRow(0) != null) 
  { 
   totalCells = sheet.getRow(0).getPhysicalNumberOfCells(); 
  } 
    
  /** ѭ��Excel���� */
  for (int r = 0; r < totalRows; r++) 
  { 
   Row row = sheet.getRow(r); 
   if (row == null) 
   { 
    continue; 
   } 
     
   List<String> rowLst = new ArrayList<String>(); 
     
   /** ѭ��Excel���� */
   for (int c = 0; c < getTotalCells(); c++) 
   { 
      
    Cell cell = row.getCell(c); 
    String cellValue = ""; 
      
    if (null != cell) 
    { 
     // �������ж����ݵ����� 
     switch (cell.getCellType()) 
     { 
      case HSSFCell.CELL_TYPE_NUMERIC: // ���� 
       cellValue = cell.getNumericCellValue() + ""; 
       break; 
        
      case HSSFCell.CELL_TYPE_STRING: // �ַ��� 
       cellValue = cell.getStringCellValue(); 
       break; 
        
      case HSSFCell.CELL_TYPE_BOOLEAN: // Boolean 
       cellValue = cell.getBooleanCellValue() + ""; 
       break; 
        
      case HSSFCell.CELL_TYPE_FORMULA: // ��ʽ 
       cellValue = cell.getCellFormula() + ""; 
       break; 
        
      case HSSFCell.CELL_TYPE_BLANK: // ��ֵ 
       cellValue = ""; 
       break; 
        
      case HSSFCell.CELL_TYPE_ERROR: // ���� 
       cellValue = "�Ƿ��ַ�"; 
       break; 
        
      default: 
       cellValue = "δ֪����"; 
       break; 
     } 
    } 
      
    rowLst.add(cellValue); 
   } 
     
   /** �����r�еĵ�c�� */
   dataLst.add(rowLst); 
  } 
    
  return dataLst; 
 } 
   
 /** 
  * 
  * ��ָ�������ȡʵ������ 
  * <��˳��������ע���ʵ���Ա������> 
  * 
  * @param wb ������ 
  * @param t ʵ�� 
  * @param in ������ 
  * @param integers ָ����Ҫ���������� 
  * @return T ��Ӧʵ�� 
  * @throws IOException 
  * @throws Exception 
  * @see [�ࡢ��#��������#��Ա] 
  */
 @SuppressWarnings("unused") 
 public static <T> T readDateT(Workbook wb, T t, InputStream in, Integer[]... integers) 
  throws IOException, Exception 
 { 
  // ��ȡ�ù������еĵ�һ�������� 
  Sheet sheet = wb.getSheetAt(0); 
    
  // ��Ա������ֵ 
  Object entityMemberValue = ""; 
    
  // ���г�Ա���� 
  Field[] fields = t.getClass().getDeclaredFields(); 
  // �п�ʼ�±� 
  int startCell = 0; 
    
  /** ѭ������Ҫ�ĳ�Ա */
  for (int f = 0; f < fields.length; f++) 
  { 
     
   fields[f].setAccessible(true); 
   String fieldName = fields[f].getName(); 
   boolean fieldHasAnno = fields[f].isAnnotationPresent(IsNeeded.class); 
   // ��ע�� 
   if (fieldHasAnno) 
   { 
    IsNeeded annotation = fields[f].getAnnotation(IsNeeded.class); 
    boolean isNeeded = annotation.isNeeded(); 
      
    // Excel��Ҫ��ֵ���� 
    if (isNeeded) 
    { 
       
     // ��ȡ�к��� 
     int x = integers[startCell][0] - 1; 
     int y = integers[startCell][1] - 1; 
       
     Row row = sheet.getRow(x); 
     Cell cell = row.getCell(y); 
       
     if (row == null) 
     { 
      continue; 
     } 
       
     // Excel�н�����ֵ 
     String cellValue = getCellValue(cell); 
     // ��Ҫ������Ա������ֵ 
     entityMemberValue = getEntityMemberValue(entityMemberValue, fields, f, cellValue); 
     // ��ֵ 
     PropertyUtils.setProperty(t, fieldName, entityMemberValue); 
     // �е��±��1 
     startCell++; 
    } 
   } 
     
  } 
    
  return t; 
 } 
   
 /** 
  * 
  * ��ȡ�б����� 
  * <��˳��������ע���ʵ���Ա������> 
  * 
  * @param wb ������ 
  * @param t ʵ�� 
  * @param beginLine ��ʼ���� 
  * @param totalcut ����������ȥ��Ӧ���� 
  * @return List<T> ʵ���б� 
  * @throws Exception 
  * @see [�ࡢ��#��������#��Ա] 
  */
 @SuppressWarnings("unchecked") 
 public static <T> List<T> readDateListT(Workbook wb, T t, int beginLine, int totalcut) 
  throws Exception 
 { 
  List<T> listt = new ArrayList<T>(); 
    
  /** �õ���һ��shell */
  Sheet sheet = wb.getSheetAt(0); 
    
  /** �õ�Excel������ */
  totalRows = sheet.getPhysicalNumberOfRows(); 
    
  /** �õ�Excel������ */
  if (totalRows >= 1 && sheet.getRow(0) != null) 
  { 
   totalCells = sheet.getRow(0).getPhysicalNumberOfCells(); 
  } 
    
  /** ѭ��Excel���� */
  for (int r = beginLine - 1; r < totalRows - totalcut; r++) 
  { 
   Object newInstance = t.getClass().newInstance(); 
   Row row = sheet.getRow(r); 
   if (row == null) 
   { 
    continue; 
   } 
     
   // ��Ա������ֵ 
   Object entityMemberValue = ""; 
     
   // ���г�Ա���� 
   Field[] fields = t.getClass().getDeclaredFields(); 
   // �п�ʼ�±� 
   int startCell = 0; 
     
   for (int f = 0; f < fields.length; f++) 
   { 
      
    fields[f].setAccessible(true); 
    String fieldName = fields[f].getName(); 
    boolean fieldHasAnno = fields[f].isAnnotationPresent(IsNeeded.class); 
    // ��ע�� 
    if (fieldHasAnno) 
    { 
     IsNeeded annotation = fields[f].getAnnotation(IsNeeded.class); 
     boolean isNeeded = annotation.isNeeded(); 
     // Excel��Ҫ��ֵ���� 
     if (isNeeded) 
     { 
      Cell cell = row.getCell(startCell); 
      String cellValue = getCellValue(cell); 
      entityMemberValue = getEntityMemberValue(entityMemberValue, fields, f, cellValue); 
      // ��ֵ 
      PropertyUtils.setProperty(newInstance, fieldName, entityMemberValue); 
      // �е��±��1 
      startCell++; 
     } 
    } 
      
   } 
     
   listt.add((T)newInstance); 
  } 
    
  return listt; 
 } 
   
 /** 
  * 
  * ����Excel����е������ж����͵õ�ֵ 
  * 
  * @param cell 
  * @return 
  * @see [�ࡢ��#��������#��Ա] 
  */
 private static String getCellValue(Cell cell) 
 { 
  String cellValue = ""; 
    
  if (null != cell) 
  { 
   // �������ж����ݵ����� 
   switch (cell.getCellType()) 
   { 
    case HSSFCell.CELL_TYPE_NUMERIC: // ���� 
     if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) 
     { 
      Date theDate = cell.getDateCellValue(); 
      SimpleDateFormat dff = new SimpleDateFormat("yyyy-MM-dd"); 
      cellValue = dff.format(theDate); 
     } 
     else
     { 
      DecimalFormat df = new DecimalFormat("0"); 
      cellValue = df.format(cell.getNumericCellValue()); 
     } 
     break; 
    case HSSFCell.CELL_TYPE_STRING: // �ַ��� 
     cellValue = cell.getStringCellValue(); 
     break; 
      
    case HSSFCell.CELL_TYPE_BOOLEAN: // Boolean 
     cellValue = cell.getBooleanCellValue() + ""; 
     break; 
      
    case HSSFCell.CELL_TYPE_FORMULA: // ��ʽ 
     cellValue = cell.getCellFormula() + ""; 
     break; 
      
    case HSSFCell.CELL_TYPE_BLANK: // ��ֵ 
     cellValue = ""; 
     break; 
      
    case HSSFCell.CELL_TYPE_ERROR: // ���� 
     cellValue = "�Ƿ��ַ�"; 
     break; 
      
    default: 
     cellValue = "δ֪����"; 
     break; 
   } 
     
  } 
  return cellValue; 
 } 
   
 /** 
  * 
  * ����ʵ���Ա���������͵õ���Ա������ֵ 
  * 
  * @param realValue 
  * @param fields 
  * @param f 
  * @param cellValue 
  * @return 
  * @see [�ࡢ��#��������#��Ա] 
  */
 private static Object getEntityMemberValue(Object realValue, Field[] fields, int f, String cellValue) 
 { 
  String type = fields[f].getType().getName(); 
  switch (type) 
  { 
   case "char": 
   case "java.lang.Character": 
   case "java.lang.String": 
    realValue = cellValue; 
    break; 
   case "java.util.Date": 
    realValue = StringUtils.isBlank(cellValue) ? null : DateUtil.strToDate(cellValue, DateUtil.YYYY_MM_DD); 
    break; 
   case "java.lang.Integer": 
    realValue = StringUtils.isBlank(cellValue) ? null : Integer.valueOf(cellValue); 
    break; 
   case "int": 
   case "float": 
   case "double": 
   case "java.lang.Double": 
   case "java.lang.Float": 
   case "java.lang.Long": 
   case "java.lang.Short": 
   case "java.math.BigDecimal": 
    realValue = StringUtils.isBlank(cellValue) ? null : new BigDecimal(cellValue); 
    break; 
   default: 
    break; 
  } 
  return realValue; 
 } 
   
 /** 
  * 
  * ����·�����ļ���ѡ��Excel�汾 
  * 
  * 
  * @param filePathOrName 
  * @param in 
  * @return 
  * @throws IOException 
  * @see [�ࡢ��#��������#��Ա] 
  */
 public static Workbook chooseWorkbook(String filePathOrName, InputStream in) 
  throws IOException 
 { 
  /** ���ݰ汾ѡ�񴴽�Workbook�ķ�ʽ */
  Workbook wb = null; 
  boolean isExcel2003 = ExcelVersionUtil.isExcel2003(filePathOrName); 
    
  if (isExcel2003) 
  { 
   wb = new HSSFWorkbook(in); 
  } 
  else
  { 
   wb = new XSSFWorkbook(in); 
  } 
    
  return wb; 
 } 
   
 static class ExcelVersionUtil 
 { 
    
  /** 
   * 
   * �Ƿ���2003��excel������true��2003 
   * 
   * 
   * @param filePath 
   * @return 
   * @see [�ࡢ��#��������#��Ա] 
   */
  public static boolean isExcel2003(String filePath) 
  { 
   return filePath.matches("^.+\\.(?i)(xls)$"); 
     
  } 
    
  /** 
   * 
   * �Ƿ���2007��excel������true��2007 
   * 
   * 
   * @param filePath 
   * @return 
   * @see [�ࡢ��#��������#��Ա] 
   */
  public static boolean isExcel2007(String filePath) 
  { 
   return filePath.matches("^.+\\.(?i)(xlsx)$"); 
     
  } 
    
 } 
   
 public static class DateUtil 
 { 
    
  // ======================���ڸ�ʽ������=====================// 
    
  public static final String YYYY_MM_DDHHMMSS = "yyyy-MM-dd HH:mm:ss"; 
    
  public static final String YYYY_MM_DD = "yyyy-MM-dd"; 
    
  public static final String YYYY_MM = "yyyy-MM"; 
    
  public static final String YYYY = "yyyy"; 
    
  public static final String YYYYMMDDHHMMSS = "yyyyMMddHHmmss"; 
    
  public static final String YYYYMMDD = "yyyyMMdd"; 
    
  public static final String YYYYMM = "yyyyMM"; 
    
  public static final String YYYYMMDDHHMMSS_1 = "yyyy/MM/dd HH:mm:ss"; 
    
  public static final String YYYY_MM_DD_1 = "yyyy/MM/dd"; 
    
  public static final String YYYY_MM_1 = "yyyy/MM"; 
    
  /** 
   * 
   * �Զ���ȡֵ��Date����תΪString���� 
   * 
   * @param date ���� 
   * @param pattern ��ʽ������ 
   * @return 
   * @see [�ࡢ��#��������#��Ա] 
   */
  public static String dateToStr(Date date, String pattern) 
  { 
   SimpleDateFormat format = null; 
     
   if (null == date) 
    return null; 
   format = new SimpleDateFormat(pattern, Locale.getDefault()); 
     
   return format.format(date); 
  } 
    
  /** 
   * ���ַ���ת����Date���͵�ʱ�� 
   * <hr> 
   * 
   * @param s �������͵��ַ���<br> 
   *   datePattern :YYYY_MM_DD<br> 
   * @return java.util.Date 
   */
  public static Date strToDate(String s, String pattern) 
  { 
   if (s == null) 
   { 
    return null; 
   } 
   Date date = null; 
   SimpleDateFormat sdf = new SimpleDateFormat(pattern); 
   try
   { 
    date = sdf.parse(s); 
   } 
   catch (ParseException e) 
   { 
    e.printStackTrace(); 
   } 
   return date; 
  } 
 } 
   
}
