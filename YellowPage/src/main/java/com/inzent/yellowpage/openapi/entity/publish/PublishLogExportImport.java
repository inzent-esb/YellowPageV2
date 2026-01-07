package com.inzent.yellowpage.openapi.entity.publish ;

import java.io.FileInputStream;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.sql.Timestamp;
import java.text.DecimalFormat;
import java.util.List ;

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
import org.springframework.stereotype.Component ;
import org.springframework.web.multipart.MultipartFile ;

import com.fasterxml.jackson.core.JsonEncoding;
import com.inzent.imanager.message.MessageGenerator;
import com.inzent.yellowpage.controller.EntityExportImportBean ;
import com.inzent.yellowpage.model.PublishLog;

@Component
public class PublishLogExportImport implements EntityExportImportBean<PublishLog>
{
  @Override
  public void exportList(HttpServletRequest request, HttpServletResponse response, PublishLog entity, List<PublishLog> list) throws Exception
  {
    String fileName = "PublishLog_" + FastDateFormat.getInstance("yyyyMMdd_hhmmss").format(new Timestamp(System.currentTimeMillis())) + ".xlsx" ;

    response.addHeader("Cache-Control", "no-cache, no-store, must-revalidate") ;
    response.addHeader("Access-Control-Expose-Headers", "Content-Disposition");
    response.setHeader("Content-Disposition", "attachment; filename=\"" + fileName + "\"; filename*=UTF-8''" + URLEncoder.encode(fileName, JsonEncoding.UTF8.getJavaName()).replaceAll("\\+", "%20")) ;
    response.setContentType("application/octet-stream") ;

    generateDownload(response, request.getServletContext().getRealPath("/template/List_PublishLog.xlsx"), entity, list) ;

    response.flushBuffer() ;
  }

  @Override
  public void exportObject(HttpServletRequest request, HttpServletResponse response, PublishLog entity) throws Exception
  {
    throw new UnsupportedOperationException() ;
  }

  @Override
  public PublishLog importObject(MultipartFile multipartFile) throws Exception
  {
    throw new UnsupportedOperationException() ;
  }

	  protected void generateDownload(HttpServletResponse response, String templateFile, PublishLog entity, List<PublishLog> list) throws Exception {
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
		      
		      // from
		      values = String.valueOf(entity.getFromDateTime().toString().substring(0,  19));
		      row = writeSheet.getRow(3);
		      cell = row.createCell(1);
		      cell.setCellStyle(cellStyle_Base);
		      cell.setCellValue(values);
		     
		      // to
		      values = String.valueOf(entity.getToDateTime().toString().substring(0,  19));
		      row = writeSheet.getRow(3);
		      cell = row.createCell(5);
		      cell.setCellStyle(cellStyle_Base);
		      cell.setCellValue(values);
		      
		      // 배포 종류
		      values = String.valueOf(entity.getDateType());
		      row = writeSheet.getRow(4);
		      cell = row.createCell(1);
		      cell.setCellStyle(cellStyle_Base);
		      cell.setCellValue(values.equals("P") ? MessageGenerator.getMessage("label.publish", "Publish") : values.equals("M") ? MessageGenerator.getMessage("label.migration", "Migration") : "");
		      
		      // 자원 종류
		      values = String.valueOf(entity.getResourceType());
		      row = writeSheet.getRow(4);
		      cell = row.createCell(3);
		      cell.setCellStyle(cellStyle_Base);
		      cell.setCellValue(values.equals("R") ? MessageGenerator.getMessage("label.modelRecord", "Record") : values.equals("S") ? MessageGenerator.getMessage("label.service", "Service") : values.equals("I") ? MessageGenerator.getMessage("label.interface", "Interface") : "" );
		      
		      // 자원 ID
		      values = entity.getResourceId();
		      row = writeSheet.getRow(4);
		      cell = row.createCell(5);
		      cell.setCellStyle(cellStyle_Base);
		      cell.setCellValue(values);
		      
		      // 상태
		      values = String.valueOf(entity.getPublishStatus());
		      row = writeSheet.getRow(5);
		      cell = row.createCell(1);
		      cell.setCellStyle(cellStyle_Base);
		      cell.setCellValue(values.equals("A") ? MessageGenerator.getMessage("label.publishLog.status.active", "Active") : values.equals("D") ? MessageGenerator.getMessage("label.publishLog.status.done", "Done") : values.equals("F") ? MessageGenerator.getMessage("label.publishLog.status.fail", "Fail") : "");
		      
		      // 이관 상태
		      values = String.valueOf(entity.getMigrationStatus());
		      row = writeSheet.getRow(5);
		      cell = row.createCell(3);
		      cell.setCellStyle(cellStyle_Base);
		      cell.setCellValue(values.equals("I") ? MessageGenerator.getMessage("label.migration.init", "Init") : values.equals("R") ? MessageGenerator.getMessage("label.migration.ready", "Ready") : values.equals("D") ? MessageGenerator.getMessage("label.migration.done", "Done") : values.equals("C") ? MessageGenerator.getMessage("label.migration.cancel", "Cancel") : "");
		      
		      // 배포자
		      values = entity.getPublishUserId();
		      row = writeSheet.getRow(5);
		      cell = row.createCell(5);
		      cell.setCellStyle(cellStyle_Base);
		      cell.setCellValue(values);
		      
		      // 메시지
		      values = entity.getPublishMessage();
		      row = writeSheet.getRow(5);
		      cell = row.createCell(7);
		      cell.setCellStyle(cellStyle_Base);
		      cell.setCellValue(values);

		      // 조회리스트 입력
		      long sum = 0;
		      int i = 7;
		          		
		      for (PublishLog publishLog : list) {
		    	  row = writeSheet.createRow(i);
		    	  
		    	  int c = 0;

		    	  // 배포일
		    	  values = null != publishLog.getPk().getPublishDateTime() ? publishLog.getPk().getPublishDateTime().substring(0, 19) : publishLog.getPk().getPublishDateTime();
		    	  cell = row.createCell(c);
		    	  cell.setCellStyle(cellStyle_Base);
		    	  cell.setCellValue(values);
		    	  
		    	  // 자원 종류
		    	  values = String.valueOf(publishLog.getResourceType());
		    	  cell = row.createCell(++c);
		    	  cell.setCellStyle(cellStyle_Base);
		    	  cell.setCellValue(values.equals("R") ? MessageGenerator.getMessage("label.modelRecord", "Record") : values.equals("S") ? MessageGenerator.getMessage("label.service", "Service") : values.equals("I") ? MessageGenerator.getMessage("label.interface", "Interface") : "" );
		    	  		    	  
		    	  // 자원 ID
		    	  values = publishLog.getResourceId();
		    	  cell = row.createCell(++c);
		    	  cell.setCellStyle(cellStyle_Base);
		    	  cell.setCellValue(values);
		    	  writeSheet.addMergedRegion(new CellRangeAddress(i, i, c, ++c));
		    	  
		    	  // 자원 버전
		    	  values = String.valueOf(publishLog.getResourceVersion());
		    	  cell = row.createCell(++c);
		    	  cell.setCellStyle(cellStyle_Base);
		    	  cell.setCellValue(values);
		    	  
		    	  // 배포 상태
		    	  values = String.valueOf(publishLog.getPublishStatus());
		    	  cell = row.createCell(++c);
		    	  cell.setCellStyle(cellStyle_Base);
		    	  cell.setCellValue(values.equals("A") ? MessageGenerator.getMessage("label.publishLog.status.active", "Active") : values.equals("D") ? MessageGenerator.getMessage("label.publishLog.status.done", "Done") : values.equals("F") ? MessageGenerator.getMessage("label.publishLog.status.fail", "Fail") : "");
		    	  
		    	  // 메시지
		    	  values = publishLog.getPublishMessage();
		    	  cell = row.createCell(++c);
		    	  cell.setCellStyle(cellStyle_Base);
		    	  cell.setCellValue(values);
		    	  
		    	  // 배포자
		    	  values = publishLog.getPublishUserId();
		    	  cell = row.createCell(++c);
		    	  cell.setCellStyle(cellStyle_Base);
		    	  cell.setCellValue(values);
		    	  
		    	  // 이관 상태
		    	  values = String.valueOf(publishLog.getMigrationStatus());
		    	  cell = row.createCell(++c);
		    	  cell.setCellStyle(cellStyle_Base);
		    	  cell.setCellValue(values.equals("I") ? MessageGenerator.getMessage("label.migration.init", "Init") : values.equals("R") ? MessageGenerator.getMessage("label.migration.ready", "Ready") : values.equals("D") ? MessageGenerator.getMessage("label.migration.done", "Done") : values.equals("C") ? MessageGenerator.getMessage("label.migration.cancel", "Cancel") : "");
		    	  
		    	  // 이관 시간
		    	  values = null != publishLog.getMigrationDateTime() ? publishLog.getMigrationDateTime().substring(0, 19) : publishLog.getMigrationDateTime();
		    	  cell = row.createCell(++c);
		    	  cell.setCellStyle(cellStyle_Base);
		    	  cell.setCellValue(values);		    	  
		    	  
		    	  sum++;
		    	  i++;
	    	  }
		      		      
		      // 총 건수
		      row = writeSheet.createRow(i);
		      DecimalFormat decFormat = new DecimalFormat("###,###");
		      values = MessageGenerator.getMessage("label.totalCount", "Total Count", decFormat.format(sum));
		      cell = row.createCell(0);
		      cell.setCellStyle(cellStyle_Info);
		      cell.setCellValue(values);
		      writeSheet.addMergedRegion(new CellRangeAddress(i, i, 0, 7));
		     		      
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
