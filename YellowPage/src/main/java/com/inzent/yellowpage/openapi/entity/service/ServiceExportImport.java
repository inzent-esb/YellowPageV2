package com.inzent.yellowpage.openapi.entity.service ;

import java.io.FileInputStream ;
import java.io.OutputStream;
import java.net.URLEncoder ;
import java.sql.Timestamp;
import java.text.DecimalFormat;
import java.util.HashMap;
import java.util.List ;
import java.util.Map;

import jakarta.servlet.http.HttpServletRequest ;
import jakarta.servlet.http.HttpServletResponse ;

import org.apache.commons.lang3.time.FastDateFormat;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell ;
import org.apache.poi.ss.usermodel.CellStyle ;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font ;
import org.apache.poi.ss.usermodel.HorizontalAlignment ;
import org.apache.poi.ss.usermodel.IndexedColors ;
import org.apache.poi.ss.usermodel.Row ;
import org.apache.poi.ss.usermodel.Sheet ;
import org.apache.poi.ss.usermodel.VerticalAlignment ;
import org.apache.poi.ss.usermodel.Workbook ;
import org.apache.poi.ss.usermodel.WorkbookFactory ;
import org.apache.poi.xssf.usermodel.XSSFCellStyle ;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component ;
import org.springframework.web.multipart.MultipartFile ;

import com.fasterxml.jackson.core.JsonEncoding ;
import com.inzent.imanager.message.MessageGenerator ;
import com.inzent.imanager.openapi.property.PropertyService ;
import com.inzent.imanager.repository.meta.Property;
import com.inzent.yellowpage.controller.EntityExportImportBean ;
import com.inzent.yellowpage.model.PublishModel;
import com.inzent.yellowpage.model.ServiceMeta ;
import com.inzent.yellowpage.openapi.entity.record.RecordExportImport ;

@Component
public class ServiceExportImport implements EntityExportImportBean<ServiceMeta>
{
  @Autowired
  protected PropertyService propertyService ;

  @Autowired
  protected ServiceService serviceService ;

  @Autowired
  protected RecordExportImport recordExportImportBean ;

  @Override
  public void exportList(HttpServletRequest request, HttpServletResponse response, ServiceMeta entity, List<ServiceMeta> list) throws Exception
  {
    String fileName = "ServiceList_" + FastDateFormat.getInstance("yyyyMMdd_hhmmss").format(new Timestamp(System.currentTimeMillis())) + ".xlsx";
    
    response.addHeader("Cache-Control", "no-cache, no-store, must-revalidate");
    response.addHeader("Access-Control-Expose-Headers", "Content-Disposition");
    response.setHeader("Content-Disposition", "attachment; filename=\"" + fileName + "\"; filename*=UTF-8''" + URLEncoder.encode(fileName, JsonEncoding.UTF8.getJavaName()).replaceAll("\\+", "%20"));
	response.setContentType("application/octet-stream");
	
	generateDownload(response, request.getServletContext().getRealPath("/template/List_Service.xlsx"), entity, list);

	response.flushBuffer();	    
  }

  @Override
  public void exportObject(HttpServletRequest request, HttpServletResponse response, ServiceMeta entity) throws Exception
  {
    try(FileInputStream fileInputStream = new FileInputStream(request.getServletContext().getRealPath("/template/ServiceTemplate.xlsx"));
    	Workbook workbook = WorkbookFactory.create(fileInputStream);
    	OutputStream outputStream = response.getOutputStream();)
    {
      String fileName = "Service_"+ entity.getId() + "_"+ FastDateFormat.getInstance("yyyyMMdd_hhmmss").format(new Timestamp(System.currentTimeMillis())) + ".xlsx";
      
      response.addHeader("Cache-Control", "no-cache, no-store, must-revalidate");
      response.addHeader("Access-Control-Expose-Headers", "Content-Disposition");
      response.setHeader("Content-Disposition", "attachment; filename=\"" + fileName + "\"; filename*=UTF-8''" + URLEncoder.encode(fileName, JsonEncoding.UTF8.getJavaName()).replaceAll("\\+", "%20"));
      response.setContentType("application/octet-stream"); 
      
      exportExcelSheet(workbook, 0, entity);
      
      if(null != entity.getRequestRecordId())
    	  recordExportImportBean.exportExcelSheet(workbook, 1, entity.getRequestRecordObject());
      
      if(null != entity.getResponseRecordId()) {
    	  boolean isSameObject = null != entity.getRequestRecordId() && entity.getRequestRecordId().equals(entity.getResponseRecordId());
    	  recordExportImportBean.exportExcelSheet(workbook, 2, isSameObject? entity.getRequestRecordObject() : entity.getResponseRecordObject());
      }

      workbook.write(outputStream) ;
    }
    catch (Exception e) 
    {
      throw e ;
	}
  }
  
  public void exportExcelSheet(Workbook workbook, int sheetIdx, ServiceMeta entity)
  {
	  Sheet writeSheet = workbook.getSheetAt(sheetIdx);
	  Row row = null;
	  Cell cell = null;

	  CellStyle cellStyle_Base = getBaseCellStyle(workbook);

	  row = writeSheet.getRow(3);

	  // ID
	  cell = row.getCell(1);
	  cell.setCellStyle(cellStyle_Base);
	  cell.setCellValue(entity.getId());

	  // 이름
	  cell = row.getCell(4);
	  cell.setCellStyle(cellStyle_Base);
	  cell.setCellValue(entity.getName());
	  
	  // 권한
	  cell = row.getCell(7);
	  cell.setCellStyle(cellStyle_Base);
	  cell.setCellValue(entity.getPrivilegeId());
	  
	  // Private
	  cell = row.getCell(9);
	  cell.setCellStyle(cellStyle_Base);
	  cell.setCellValue(String.valueOf(entity.getPrivateYn()));
		  
	  // 업무명
	  cell = row.getCell(11);
	  cell.setCellStyle(cellStyle_Base);
	  cell.setCellValue(entity.getBizId());
	  
	  row = writeSheet.getRow(4);
	  
	  // 연계 서버 ID
	  cell = row.getCell(1);
	  cell.setCellStyle(cellStyle_Base);
	  cell.setCellValue(entity.getServerId());
	  
	  // 인터페이스 타입
	  cell = row.getCell(3);
	  cell.setCellStyle(cellStyle_Base);
	  cell.setCellValue(MessageGenerator.getMessage(getInterfaceTypeMap().get(entity.getInterfaceType()), ""));
	  
	  // Target 시스템 ID
	  cell = row.getCell(5);
	  cell.setCellStyle(cellStyle_Base);
	  cell.setCellValue(entity.getTargetSystemId());
	  
	  // Meta Domain
	  cell = row.getCell(7);
	  cell.setCellStyle(cellStyle_Base);
	  cell.setCellValue(entity.getMetaDomain());
	  
	  // 요청 데이터 모델
	  cell = row.getCell(9);
	  cell.setCellStyle(cellStyle_Base);
	  cell.setCellValue(entity.getRequestRecordId());
	  
	  // 응답 데이터 모델
	  cell = row.getCell(11);
	  cell.setCellStyle(cellStyle_Base);
	  cell.setCellValue(entity.getResponseRecordId());
	  
	  row = writeSheet.getRow(5);
	  
	  // 사용 여부
	  cell = row.getCell(1);
	  cell.setCellStyle(cellStyle_Base);
	  cell.setCellValue(String.valueOf(entity.getUseYn()));
	  
	  // 작성자
	  cell = row.getCell(3);
	  cell.setCellStyle(cellStyle_Base);
	  cell.setCellValue(entity.getUpdateUserId());
	  
	  // 작성일
	  cell = row.getCell(5);
	  cell.setCellStyle(cellStyle_Base);
	  cell.setCellValue(String.valueOf(entity.getUpdateTimestamp()));
	  
	  row = writeSheet.getRow(6);
	  
	  // 설명
	  cell = row.getCell(1);
	  cell.setCellStyle(cellStyle_Base);
	  cell.setCellValue(entity.getDescription());
  }

  @Override
  public ServiceMeta importObject(MultipartFile multipartFile) throws Exception
  {
    ServiceMeta serviceMeta = new ServiceMeta() ;

    try (OPCPackage opcPackage = OPCPackage.open(multipartFile.getInputStream());
	     Workbook workbook = new XSSFWorkbook(opcPackage)) {
    	
    	serviceMeta = importExcelSheet(workbook, 0);
      
      if(null != serviceMeta.getRequestRecordId())
    	  serviceMeta.setRequestRecordObject(recordExportImportBean.importExcelSheet(workbook, 1));
      
      if(null != serviceMeta.getResponseRecordId())
    	  serviceMeta.setResponseRecordObject(recordExportImportBean.importExcelSheet(workbook, null != serviceMeta.getRequestRecordId()? 2 : 1));
      
    }
    catch (Exception e) 
    {
      throw e ;
	}

    return serviceMeta ;
  }
  
  public ServiceMeta importExcelSheet(Workbook workbook, int sheetIdx) throws Exception
  {
	  ServiceMeta serviceMeta = new ServiceMeta();
		
		Sheet sheet = workbook.getSheetAt(sheetIdx);
		Row row = null;
	    Cell cell = null;
	    
	    row = sheet.getRow(3);

	    // ID
	    cell = row.getCell(1);
	    serviceMeta.setId(cell.getStringCellValue());

	    // 이름
		cell = row.getCell(4);
		serviceMeta.setName(cell.getStringCellValue());
		  
		// 권한
		cell = row.getCell(7);
		serviceMeta.setPrivilegeId(cell.getStringCellValue());
		
		// Private
		cell = row.getCell(9);
		serviceMeta.setPrivateYn(cell.getStringCellValue().charAt(0));
			  
		// 업무명
		cell = row.getCell(11);
		serviceMeta.setBizId(cell.getStringCellValue());
		  
		row = sheet.getRow(4);
		  
		// 연계 서버 ID
		cell = row.getCell(1);
		serviceMeta.setServerId(cell.getStringCellValue());
		  
		// 인터페이스 타입
		cell = row.getCell(3);
		
	    for (Property property : propertyService.getProperties("Interface.Type", true)) {
	    	if(cell.getStringCellValue().trim().equals(MessageGenerator.getMessage(property.getPropertyValue(), "").trim())) {
		    	serviceMeta.setInterfaceType(property.getPk().getPropertyKey());
		    	break;
	    	}
	    }
		
		// Target 시스템 ID
		cell = row.getCell(5);
		serviceMeta.setTargetSystemId(cell.getStringCellValue());
		  
		// Meta Domain
		cell = row.getCell(7);
		serviceMeta.setMetaDomain(cell.getStringCellValue());
		  
		// 요청 데이터 모델
		cell = row.getCell(9);
		serviceMeta.setRequestRecordId(0 == cell.getStringCellValue().length()? null : cell.getStringCellValue());
		  
		// 응답 데이터 모델
		cell = row.getCell(11);
		serviceMeta.setResponseRecordId(0 == cell.getStringCellValue().length()? null : cell.getStringCellValue());
		  
		row = sheet.getRow(5);
		  
		// 사용 여부
		cell = row.getCell(1);
		serviceMeta.setUseYn(cell.getStringCellValue().charAt(0));
		  
		// 작성자
		cell = row.getCell(3);
		serviceMeta.setUpdateUserId(cell.getStringCellValue());
		  
		row = sheet.getRow(6);
		  
		// 설명
		cell = row.getCell(1);
		serviceMeta.setDescription(cell.getStringCellValue());
		
		// 배포 상태
		serviceMeta.setPublishStatus(PublishModel.PUBLISH_MAKE);
		
		// Property
		ServiceMeta source = serviceService.get(serviceMeta.getId());		

		if (null != source) {	
			serviceMeta.setServiceProperties(source.getServiceProperties());
		} else {
			serviceService.generateProperties(serviceMeta, false);
		}
	    
		return serviceMeta;
	}
  
  protected void generateDownload(HttpServletResponse response, String templateFile, ServiceMeta entity, List<ServiceMeta> list) throws Exception
  {
    try(FileInputStream fileInputStream = new FileInputStream(templateFile);
    	Workbook workbook = WorkbookFactory.create(fileInputStream);
	    OutputStream outputStream = response.getOutputStream();)
    {
      Sheet writeSheet ;
      Row row = null ;
      Cell cell = null ;
      String values = null ;
    	    
      writeSheet = workbook.getSheetAt(0) ;
    	    
      // Cell 스타일 지정.
      CellStyle cellStyle_Base = getBaseCellStyle(workbook);
        
      // 서비스 ID
      values = entity.getId();
      row = writeSheet.getRow(3);
      cell = row.createCell(1);
      cell.setCellStyle(cellStyle_Base);
      cell.setCellValue(values);
    	
      // 서비스 이름
      values = entity.getName();
      row = writeSheet.getRow(3);
      cell = row.createCell(3);
      cell.setCellStyle(cellStyle_Base);
      cell.setCellValue(values);
    	
      // 시스템 ID
      values = entity.getTargetSystemId();
      row = writeSheet.getRow(3);
      cell = row.createCell(5);
      cell.setCellStyle(cellStyle_Base);
      cell.setCellValue(values);
    	
      // 업무명
      values = entity.getBizId();
      row = writeSheet.getRow(3);
      cell = row.createCell(7);
      cell.setCellStyle(cellStyle_Base);
      cell.setCellValue(values);
    	
      // 연계 서버 ID
      values = entity.getServerId();
      row = writeSheet.getRow(3);
      cell = row.createCell(9);
      cell.setCellStyle(cellStyle_Base);
      cell.setCellValue(values);		
        
      // 인터페이스 타입
      values = MessageGenerator.getMessage(getInterfaceTypeMap().get(entity.getInterfaceType()), "");
      row = writeSheet.getRow(4);
      cell = row.createCell(1);
      cell.setCellStyle(cellStyle_Base);
      cell.setCellValue(values);
    	
      // 배포 상태
      char publishStatus = entity.getPublishStatus();
      values = PublishModel.PUBLISH_MAKE == publishStatus? "MAKE" : 
    	  	   PublishModel.PUBLISH_REQUEST == publishStatus? "REQUEST" : 
    	  	   PublishModel.PUBLISH_CANCEL == publishStatus? "CANCEL" : 
    	  	   PublishModel.PUBLISH_APPROVE == publishStatus? "APPROVE" : "";
      row = writeSheet.getRow(4);
      cell = row.createCell(3);
      cell.setCellStyle(cellStyle_Base);
      cell.setCellValue(values);	
    	
      // 사용여부
      values = 'Y' == entity.getUseYn()? MessageGenerator.getMessage("label.yes", "yes") : 'N' == entity.getUseYn()? MessageGenerator.getMessage("label.no", "no") : "";
      row = writeSheet.getRow(4);
      cell = row.createCell(5);
      cell.setCellStyle(cellStyle_Base);
      cell.setCellValue(values);
    	
      // 비고
      values = entity.getDescription();
      row = writeSheet.getRow(4);
      cell = row.createCell(7);
      cell.setCellStyle(cellStyle_Base);
      cell.setCellValue(values);
    	
      // 조회리스트 입력
      long sum = 0;
      int i = 6;
    	
      for (ServiceMeta serviceMeta : list)
      {
        row = writeSheet.createRow(i);
    		  
        int c = 0;
    	  
        // 서비스 ID
        values = serviceMeta.getId();
        cell = row.createCell(c);
        cell.setCellStyle(cellStyle_Base);
        cell.setCellValue(values);
    	  
        // 서비스 이름
        values = serviceMeta.getName();
        cell = row.createCell(++c);
        cell.setCellStyle(cellStyle_Base);
        cell.setCellValue(values);
    	  
        // 시스템 ID
        values = serviceMeta.getTargetSystemId();
        cell = row.createCell(++c);
        cell.setCellStyle(cellStyle_Base);
        cell.setCellValue(values);
    	  
        // 업무명
        values = serviceMeta.getBizId();
        cell = row.createCell(++c);
        cell.setCellStyle(cellStyle_Base);
        cell.setCellValue(values);
    	  
        // 연계 서버 ID
        values = serviceMeta.getServerId();
        cell = row.createCell(++c);
        cell.setCellStyle(cellStyle_Base);
        cell.setCellValue(values);
        
        // 인터페이스 타입
        values = MessageGenerator.getMessage(getInterfaceTypeMap().get(serviceMeta.getInterfaceType()), "");
        cell = row.createCell(++c);
        cell.setCellStyle(cellStyle_Base);
        cell.setCellValue(values);
    	  
        // 배포 상태
        publishStatus = serviceMeta.getPublishStatus();
        values = PublishModel.PUBLISH_MAKE == publishStatus? "MAKE" : 
	 	  	   	 PublishModel.PUBLISH_REQUEST == publishStatus? "REQUEST" : 
	 	  	   	 PublishModel.PUBLISH_CANCEL == publishStatus? "CANCEL" : 
	 	  	     PublishModel.PUBLISH_APPROVE == publishStatus? "APPROVE" : "";
        cell = row.createCell(++c);
        cell.setCellStyle(cellStyle_Base);
        cell.setCellValue(values);
    	  
        // 배포 요청 시간
        values = serviceMeta.getPublishDateTime();
        cell = row.createCell(++c);
        cell.setCellStyle(cellStyle_Base);
        cell.setCellValue(values);	  
    	  
        // 배포 완료 시간
        values = null != serviceMeta.getPublishTimestamp()? serviceMeta.getPublishTimestamp().toString() : null;
        cell = row.createCell(++c);
        cell.setCellStyle(cellStyle_Base);
        cell.setCellValue(values);
    	  
        // 사용 여부
        values = 'Y' == serviceMeta.getUseYn()? MessageGenerator.getMessage("label.yes", "yes") : 'N' == serviceMeta.getUseYn()? MessageGenerator.getMessage("label.no", "no") : "";
        cell = row.createCell(++c);
        cell.setCellStyle(cellStyle_Base);
        cell.setCellValue(values);
    	  
        // Version
        values = String.valueOf(serviceMeta.getUpdateVersion());
        cell = row.createCell(++c);
        cell.setCellStyle(cellStyle_Base);
        cell.setCellValue(values);
    	  
        // 작성자
        values = serviceMeta.getUpdateUserId();
        cell = row.createCell(++c);
        cell.setCellStyle(cellStyle_Base);
        cell.setCellValue(values);
    	  
        // 작성일
        values = null != serviceMeta.getUpdateTimestamp()? serviceMeta.getUpdateTimestamp().toString().substring(0, 19) : "";
        cell = row.createCell(++c);
        cell.setCellStyle(cellStyle_Base);
        cell.setCellValue(values);
    	  
        sum++;
        i++;		  
      }
    	
      // 합계
      row = writeSheet.createRow(i);
      DecimalFormat decFormat = new DecimalFormat("###,###");
      values = MessageGenerator.getMessage("label.totalCount", "Total", decFormat.format(sum));
      cell = row.createCell(0);
      cell.setCellStyle(cellStyle_Base);
      cell.setCellValue(values);
    	
      list = null ;
      workbook.write(outputStream);    	
    }
    catch (Exception e) 
    {
      throw e ;
	}
  }

  public XSSFCellStyle getBaseCellStyle(Workbook workbook) 
  {
    // Cell 스타일 지정.
	XSSFCellStyle cellStyle = (XSSFCellStyle) workbook.createCellStyle();
	// 텍스트 맞춤(세로가운데)
	cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
	// 텍스트 맞춤 (가로 가운데)
	cellStyle.setAlignment(HorizontalAlignment.CENTER);

	// 폰트 지정 사이즈 10
	cellStyle.setFont(getBaseFont(workbook, 10, IndexedColors.BLACK.getIndex()));
	
	// Cell 테두리 (선)
	cellStyle.setBorderBottom(BorderStyle.THIN);
	cellStyle.setBorderLeft(BorderStyle.THIN);
	cellStyle.setBorderRight(BorderStyle.THIN);
	cellStyle.setBorderTop(BorderStyle.THIN);

	// Cell 잠금
	cellStyle.setLocked(true);
	// Cell 에서 Text 줄바꿈 활성화
	cellStyle.setWrapText(true);

	return cellStyle;
  }

  public XSSFCellStyle getInfoCellStyle(Workbook workbook) 
  {
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
  
  public Font getBaseFont(Workbook workbook, int size, short color)
  {
    Font font = workbook.createFont() ;
    font.setFontHeight((short) (20 * size)) ;
    font.setFontName("굴림") ;
    font.setColor(color) ;
    return font ;
  }
  
  protected Map<String, String> getInterfaceTypeMap()
  {
    Map<String, String> map = new HashMap<>() ;

    for (Property property : propertyService.getProperties("Interface.Type", true))
      map.put(property.getPk().getPropertyKey(), property.getPropertyValue()) ;

    return map ;
  }
}
