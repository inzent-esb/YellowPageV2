package com.inzent.yellowpage.openapi.statistics;

import java.io.FileInputStream;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.sql.Timestamp;
import java.text.DecimalFormat;
import java.util.List ;
import java.util.Map;
import java.util.stream.Collectors ;

import jakarta.servlet.http.HttpServletRequest ;
import jakarta.servlet.http.HttpServletResponse ;

import org.apache.commons.lang3.time.FastDateFormat;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.springframework.beans.factory.annotation.Autowired ;
import org.springframework.stereotype.Component ;
import org.springframework.web.multipart.MultipartFile ;

import com.fasterxml.jackson.core.JsonEncoding;
import com.inzent.imanager.message.MessageGenerator;
import com.inzent.imanager.openapi.property.PropertyService ;
import com.inzent.yellowpage.controller.EntityExportImportBean ;
import com.inzent.yellowpage.model.StatsUser;

@Component
public class StatsUserExportImport implements EntityExportImportBean<StatsUser>
{
  @Autowired
  protected PropertyService propertyService ;

  @Override
  public void exportList(HttpServletRequest request, HttpServletResponse response, StatsUser entity, List<StatsUser> list) throws Exception
  {
    String fileName = "LogsStatsUser_" + FastDateFormat.getInstance("yyyyMMdd_hhmmss").format(new Timestamp(System.currentTimeMillis())) + ".xlsx" ;

    response.addHeader("Cache-Control", "no-cache, no-store, must-revalidate") ;
    response.addHeader("Access-Control-Expose-Headers", "Content-Disposition");
    response.setHeader("Content-Disposition", "attachment; filename=\"" + fileName + "\"; filename*=UTF-8''" + URLEncoder.encode(fileName, JsonEncoding.UTF8.getJavaName()).replaceAll("\\+", "%20")) ;
    response.setContentType("application/octet-stream") ;

    generateDownload(response, request.getServletContext().getRealPath("/template/LogsStatsUser.xlsx"), getInterfaceTypeMap(), entity, list) ;

    response.flushBuffer() ;
  }

  @Override
  public void exportObject(HttpServletRequest request, HttpServletResponse response, StatsUser entity) throws Exception
  {
    throw new UnsupportedOperationException() ;
  }

  @Override
  public StatsUser importObject(MultipartFile multipartFile) throws Exception
  {
    throw new UnsupportedOperationException() ;
  }

  protected Map<String, String> getInterfaceTypeMap()
  {
    return propertyService.getProperties("Interface.Type", true).stream().collect(
        Collectors.toMap(property -> property.getPk().getPropertyKey(), property -> property.getPropertyValue())) ;
  }

  protected void generateDownload(HttpServletResponse response, String templateFile, Map<String, String> properties, StatsUser entity, List<StatsUser> list) throws Exception {
	  try (OutputStream outputStream = response.getOutputStream();
           FileInputStream fileInputStream = new FileInputStream(templateFile);
           Workbook workbook = WorkbookFactory.create(fileInputStream);) {
		  
		  Sheet writeSheet = workbook.getSheetAt(0);
	      Row row = null ;
	      Cell cell = null ;
	      String values = null ;

	      // Cell 스타일 지정
	      CellStyle cellStyle_Base = getBaseCellStyle(workbook);
	      CellStyle cellStyle_Info = getInfoCellStyle(workbook);

	      // 서버 ID
	      values = entity.getPk().getUserId();
	      row = writeSheet.getRow(3);
	      cell = row.createCell(1);
	      cell.setCellStyle(cellStyle_Base);
	      cell.setCellValue(values);
	      
	      // 업무 ID
	      values = entity.getPk().getBizId();	      
	      row = writeSheet.getRow(3);
	      cell = row.createCell(3);
	      cell.setCellStyle(cellStyle_Base);
	      cell.setCellValue(values);
	      
	      // 구분
	      if('I' == entity.getPk().getStatsType())			values = MessageGenerator.getMessage("label.interface", "Interface");
    	  else if ('S' == entity.getPk().getStatsType())	values = MessageGenerator.getMessage("label.service", "Service"); 	     
    	  else 												values = "";
	      row = writeSheet.getRow(3);
	      cell = row.createCell(5);
	      cell.setCellStyle(cellStyle_Base);
	      cell.setCellValue(values);
	      
	      // 인터페이스 타입
	      values = MessageGenerator.getMessage(properties.get(entity.getPk().getInterfaceType()), "");   
	      row = writeSheet.getRow(4);
	      cell = row.createCell(1);
	      cell.setCellStyle(cellStyle_Base);
	      cell.setCellValue(values);	      
	      
	      // 조회리스트 입력
	      long sum = 0, editSum = 0, doneSum = 0, errorSum = 0;
	      int i = 6;
	          		
	      for (StatsUser statsUser : list) {
	    	  row = writeSheet.createRow(i);
	    	  
	    	  int c = 0;
	    	  
	    	  // 사용자 ID
	    	  values = statsUser.getPk().getUserId();
	    	  cell = row.createCell(c);
	    	  cell.setCellStyle(cellStyle_Base);
	    	  cell.setCellValue(values);
	        
	    	  // 업무 ID
	    	  values = statsUser.getPk().getBizId();
	    	  cell = row.createCell(++c);
	    	  cell.setCellStyle(cellStyle_Base);
	    	  cell.setCellValue(values);
	        
	    	  // 구분
	    	  if('I' == statsUser.getPk().getStatsType())			values = MessageGenerator.getMessage("label.interface", "Interface");
	    	  else if ('S' == statsUser.getPk().getStatsType())		values = MessageGenerator.getMessage("label.service", "Service");
	    	  cell = row.createCell(++c);
	    	  cell.setCellStyle(cellStyle_Base);
	    	  cell.setCellValue(values);
	        	  
	    	  // 인터페이스 타입
	    	  values = MessageGenerator.getMessage(properties.get(statsUser.getPk().getInterfaceType()), "");
	    	  cell = row.createCell(++c);
	    	  cell.setCellStyle(cellStyle_Base);
	    	  cell.setCellValue(values);	    	  
	    	  
	    	  // 인터페이스 등록 개수
	    	  values = String.valueOf(statsUser.getEditCount());
	    	  cell = row.createCell(++c);
	    	  cell.setCellStyle(cellStyle_Base);
	    	  cell.setCellValue(values);
	        
	    	  // 배포 성공 개수
	    	  values = String.valueOf(statsUser.getPublishDoneCount());
	    	  cell = row.createCell(++c);
	    	  cell.setCellStyle(cellStyle_Base);
	    	  cell.setCellValue(values);
	        
	    	  // 배포 실패 개수
	    	  values = String.valueOf(statsUser.getPublishErrorCount());
	    	  cell = row.createCell(++c);
	    	  cell.setCellStyle(cellStyle_Base);
	    	  cell.setCellValue(values);
	        
	    	  editSum += statsUser.getEditCount();
	    	  doneSum += statsUser.getPublishDoneCount();
	    	  errorSum += statsUser.getPublishErrorCount();
	    	  
	    	  sum++;
	    	  i++;
    	  }
	      
	      // 합계
	      row = writeSheet.createRow(i);
	      values = MessageGenerator.getMessage("label.total", "Total");
	      cell = row.createCell(0);
	      cell.setCellStyle(cellStyle_Info);
	      cell.setCellValue(values);
	      
	      // 총 건수
	      DecimalFormat decFormat = new DecimalFormat("###,###");
	      values = MessageGenerator.getMessage("label.totalCount", "Total Count", decFormat.format(sum));
	      cell = row.createCell(1);
	      cell.setCellStyle(cellStyle_Info);
	      cell.setCellValue(values);
	      writeSheet.addMergedRegion(new CellRangeAddress(i, i, 1, 3));
	      cell = row.createCell(3);
	      
	      // 인터페이스 등록 개수 합계
	      cell = row.createCell(4);
	      cell.setCellStyle(cellStyle_Base);
	      cell.setCellValue(editSum);
	      
	      // 배포 성공 횟수 합계
	      cell = row.createCell(5);
	      cell.setCellStyle(cellStyle_Base);
	      cell.setCellValue(doneSum);
	      
	      // 배포 실패 횟수 합계
	      cell = row.createCell(6);
	      cell.setCellStyle(cellStyle_Base);
	      cell.setCellValue(errorSum);
	      
	      list = null ;
	      workbook.write(outputStream);
	  } catch (Exception e) {
		  throw e ;
	  }
  }
    
  public XSSFCellStyle getBaseCellStyle(Workbook workbook) {
	  // Cell 스타일 지정.
	  XSSFCellStyle cellStyle = (XSSFCellStyle) workbook.createCellStyle();
	  // 텍스트 맞춤(세로 가운데)
	  cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
	  // 텍스트 맞춤 (가로 가운데)
	  cellStyle.setAlignment(HorizontalAlignment.CENTER);

	  // 폰트 지정 사이즈 10
	  cellStyle.setFont(getBaseFont(workbook, 10, IndexedColors.BLACK.getIndex()));

	  // Cell 잠금
	  cellStyle.setLocked(true);
	  // Cell 에서 Text 줄바꿈 활성화
	  cellStyle.setWrapText(true);

	  return cellStyle;
  }

  public XSSFCellStyle getInfoCellStyle(Workbook workbook) {
	  XSSFCellStyle cellStyle = getBaseCellStyle(workbook);
	  cellStyle.setAlignment(HorizontalAlignment.CENTER);

	  // 폰트 지정 사이즈 (굵게)
	  Font font = getBaseFont(workbook, 10, IndexedColors.BLACK.getIndex());
	  font.setBold(true);
	  cellStyle.setFont(font);

	  cellStyle.setFillForegroundColor(new XSSFColor(new byte[] { (byte) 242, (byte) 242, (byte) 242 }, null));
	  cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

	  return cellStyle;
  }
  
  public Font getBaseFont(Workbook workbook, int size, short color) {
	  // 폰트
	  Font font = workbook.createFont();
	  font.setFontHeight((short) (20 * size));
	  font.setFontName("굴림");
	  font.setColor(color);

	  return font;
  }
}
