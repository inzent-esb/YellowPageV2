package com.inzent.yellowpage.openapi.entity.interfaces ;

import java.io.FileInputStream;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.sql.Timestamp;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List ;
import java.util.stream.Collectors;

import jakarta.servlet.http.HttpServletRequest ;
import jakarta.servlet.http.HttpServletResponse ;

import org.apache.commons.lang3.time.FastDateFormat;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BorderStyle;
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
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier ;
import org.springframework.stereotype.Component ;
import org.springframework.web.multipart.MultipartFile ;

import com.fasterxml.jackson.core.JsonEncoding;
import com.inzent.imanager.message.MessageGenerator;
import com.inzent.imanager.service.MetaEntityService ;
import com.inzent.yellowpage.controller.EntityExportImportBean ;
import com.inzent.yellowpage.model.InterfaceClass;
import com.inzent.yellowpage.model.InterfaceMeta ;
import com.inzent.yellowpage.model.InterfaceResponse;
import com.inzent.yellowpage.model.InterfaceResponsePK;
import com.inzent.yellowpage.model.ModelRecord;
import com.inzent.yellowpage.model.PublishModel;
import com.inzent.yellowpage.openapi.entity.record.RecordExportImport ;

@Component
public class InterfaceExportImport implements EntityExportImportBean<InterfaceMeta>
{
  @Autowired
  @Qualifier("interfaceClassService")
  protected MetaEntityService<String, InterfaceClass> interfaceClassService ;

  @Autowired
  protected InterfaceService interfaceService ;

  @Autowired
  protected RecordExportImport recordExportImport ;

  @Override
  public void exportList(HttpServletRequest request, HttpServletResponse response, InterfaceMeta entity, List<InterfaceMeta> list) throws Exception
  {
    String fileName = "InterfaceList_" + FastDateFormat.getInstance("yyyyMMdd_hhmmss").format(new Timestamp(System.currentTimeMillis())) + ".xlsx";
		
    response.addHeader("Cache-Control", "no-cache, no-store, must-revalidate");
    response.addHeader("Access-Control-Expose-Headers", "Content-Disposition");
    response.setHeader("Content-Disposition", "attachment; filename=\"" + fileName + "\"; filename*=UTF-8''" + URLEncoder.encode(fileName, JsonEncoding.UTF8.getJavaName()).replaceAll("\\+", "%20"));
    response.setContentType("application/octet-stream");
		
    generateDownload(response, request.getServletContext().getRealPath("/template/List_Interface.xlsx"), entity, list);

    response.flushBuffer();	
  }

  @Override
  public void exportObject(HttpServletRequest request, HttpServletResponse response, InterfaceMeta entity) throws Exception
  {
    try(FileInputStream fileInputStream = new FileInputStream(request.getServletContext().getRealPath("/template/InterfaceTemplate.xlsx"));
    	Workbook workbook = WorkbookFactory.create(fileInputStream);
    	OutputStream outputStream = response.getOutputStream();)
    {
      String fileName = "Interface_" + entity.getId() + "_" + FastDateFormat.getInstance("yyyyMMdd_hhmmss").format(new Timestamp(System.currentTimeMillis())) + ".xlsx";
      
      response.addHeader("Cache-Control", "no-cache, no-store, must-revalidate");
      response.addHeader("Access-Control-Expose-Headers", "Content-Disposition");
      response.setHeader("Content-Disposition", "attachment; filename=\"" + fileName + "\"; filename*=UTF-8''" + URLEncoder.encode(fileName, JsonEncoding.UTF8.getJavaName()).replaceAll("\\+", "%20"));
      response.setContentType("application/octet-stream");
      
      exportExcelSheet(workbook, 0, entity);
      
      if(null != entity.getRequestRecordId())
    	  recordExportImport.exportExcelSheet(workbook, 1, entity.getRequestRecordObject());

      int responseSheetCnt = entity.getInterfaceResponses().size();
      boolean isSameObject = false;
      
      if(1 == responseSheetCnt) {
    	  isSameObject = null != entity.getRequestRecordId() && entity.getRequestRecordId().equals(entity.getInterfaceResponses().get(0).getPk().getRecordId());
		  recordExportImport.exportExcelSheet(workbook, 2, isSameObject? entity.getRequestRecordObject(): entity.getInterfaceResponses().get(0).getRecordObject()); 
      } else if(1 < responseSheetCnt) {
    	  for(int idx = 0; idx < responseSheetCnt; idx++) {
    		  int responseSheetIdx = workbook.getNumberOfSheets() - 1;
    		  
    		  workbook.cloneSheet(responseSheetIdx);
    		  workbook.setSheetName(responseSheetIdx, 0 == idx? workbook.getSheetName(2) : workbook.getSheetName(2) + (responseSheetIdx - 1));
    		  
    		  isSameObject = null != entity.getRequestRecordId() && entity.getRequestRecordId().equals(entity.getInterfaceResponses().get(0).getPk().getRecordId());
    		 
    		  recordExportImport.exportExcelSheet(workbook, responseSheetIdx, isSameObject? entity.getRequestRecordObject(): entity.getInterfaceResponses().get(idx).getRecordObject());    		  
    	  }
    	  
    	  workbook.removeSheetAt(workbook.getNumberOfSheets() - 1);
      }
      
      workbook.write(outputStream) ;
    }
    catch (Exception e) 
    {
      throw e ;
	}
  }
  
  public void exportExcelSheet(Workbook workbook, int sheetIdx, InterfaceMeta entity) throws Exception
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

	  row = writeSheet.getRow(4);
	  
	  // Source 시스템 ID
	  cell = row.getCell(1);
	  cell.setCellStyle(cellStyle_Base);
	  cell.setCellValue(entity.getSourceSystemId());
	  
	  // 업무명
	  cell = row.getCell(3);
	  cell.setCellStyle(cellStyle_Base);
	  cell.setCellValue(entity.getBizId());
	  
	  // 연계 서버 ID
	  cell = row.getCell(5);
	  cell.setCellStyle(cellStyle_Base);
	  cell.setCellValue(entity.getServerId());
	  
	  // 인터페이스 Class
	  if (null != entity.getInterfaceClass() && 0 < entity.getInterfaceClass().trim().length()) {
    	  InterfaceClass interfaceClass = interfaceClassService.search(new InterfaceClass(), -1, null, false).stream()
											    	    .filter(interfaceClassInfo -> interfaceClassInfo.getId().equals(entity.getInterfaceClass()))
											    	    .collect(Collectors.toList())
											    	    .get(0);
    	  
    	  cell = row.getCell(7);
    	  cell.setCellStyle(cellStyle_Base);
    	  cell.setCellValue(interfaceClass.getName());
      }
	  
	  // 사용 여부
	  cell = row.getCell(9);
	  cell.setCellStyle(cellStyle_Base);
	  cell.setCellValue(String.valueOf(entity.getUseYn()));
	  
	  row = writeSheet.getRow(5);
	  
	  // Meta Domain
	  cell = row.getCell(1);
	  cell.setCellStyle(cellStyle_Base);
	  cell.setCellValue(entity.getMetaDomain());
	  
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
  public InterfaceMeta importObject(MultipartFile multipartFile) throws Exception
  {
	InterfaceMeta interfaceMeta = new InterfaceMeta() ;

    try (OPCPackage opcPackage = OPCPackage.open(multipartFile.getInputStream());
   	     Workbook workbook = new XSSFWorkbook(opcPackage)) {
    
    	interfaceMeta = importExcelSheet(workbook, 0);
    
    	//요청 데이터 모델
    	if(!getIsEmptyRecordId(workbook, 1).isEmpty()) {
    		ModelRecord modelRecord = recordExportImport.importExcelSheet(workbook, 1);
    		interfaceMeta.setRequestRecordId(modelRecord.getId());
    		interfaceMeta.setRequestRecordObject(modelRecord);
    	}
    	
		//응답 데이터 모델
		List<InterfaceResponse> interfaceResponses = new ArrayList<>();
		InterfaceResponse interfaceResponse;
		InterfaceResponsePK responsePK;
		int sheetStartIdx = 2;
    	
		while(sheetStartIdx < workbook.getNumberOfSheets()) {
			if(getIsEmptyRecordId(workbook, sheetStartIdx).isEmpty()) break;
			
			ModelRecord modelRecord = recordExportImport.importExcelSheet(workbook, sheetStartIdx);
			
			responsePK = new InterfaceResponsePK();
			responsePK.setId(interfaceMeta.getId());
			responsePK.setRecordId(modelRecord.getId());
			
			interfaceResponse = new InterfaceResponse();
			interfaceResponse.setPk(responsePK);
			interfaceResponse.setRecordObject(modelRecord);
			interfaceResponse.setInterfaceMeta(interfaceMeta);    			
      		interfaceResponses.add(interfaceResponse);
      		
      		sheetStartIdx++;
		}
		
		interfaceMeta.setInterfaceResponses(interfaceResponses);
    }
    catch (Exception e) 
    {
      throw e ;
	}
    
    

    return interfaceMeta ;
  }
  
  public InterfaceMeta importExcelSheet(Workbook workbook, int sheetIdx) throws Exception
  {
	InterfaceMeta interfaceMeta = new InterfaceMeta();
	
	Sheet writeSheet = workbook.getSheetAt(sheetIdx);
	Row row = null;
	Cell cell = null;
	
	row = writeSheet.getRow(3);

	// ID
	cell = row.getCell(1);
	interfaceMeta.setId(cell.getStringCellValue());
	  
	// 이름
	cell = row.getCell(4);
	interfaceMeta.setName(cell.getStringCellValue());
	  
	// 권한
	cell = row.getCell(7);
	interfaceMeta.setPrivilegeId(cell.getStringCellValue());
	  
	// Private
	cell = row.getCell(9);	
	interfaceMeta.setPrivateYn(cell.getStringCellValue().charAt(0));

	row = writeSheet.getRow(4);
	  
	// Source 시스템 ID
	cell = row.getCell(1);
	interfaceMeta.setSourceSystemId(cell.getStringCellValue());
	  
	// 업무명
	cell = row.getCell(3);
	interfaceMeta.setBizId(cell.getStringCellValue());
	  
	// 연계 서버 ID
	cell = row.getCell(5);
	interfaceMeta.setServerId(cell.getStringCellValue());
	  
	// 인터페이스 Class
	cell = row.getCell(7);
	String value = cell.getStringCellValue();
	
	InterfaceClass interfaceClass = interfaceClassService.search(new InterfaceClass(), -1, null, false).stream()
    	    .filter(interfaceClassInfo -> value.trim().equals(interfaceClassInfo.getName().trim()))
    	    .collect(Collectors.toList())
    	    .get(0);
	
	interfaceMeta.setInterfaceClass(interfaceClass.getId());
	
	row = writeSheet.getRow(5);
	  
	// Meta Domain
	cell = row.getCell(1);
	interfaceMeta.setMetaDomain(cell.getStringCellValue());
	  
	// 작성자
	cell = row.getCell(3);
	interfaceMeta.setUpdateUserId(cell.getStringCellValue());
	  
	row = writeSheet.getRow(6);
	  
	// 설명
	cell = row.getCell(1);
	interfaceMeta.setDescription(cell.getStringCellValue());
	  
	// 배포 상태
	interfaceMeta.setPublishStatus(PublishModel.PUBLISH_MAKE);
	  
	// 사용 여부
	interfaceMeta.setUseYn('Y');
	
	InterfaceMeta source = interfaceService.get(interfaceMeta.getId());
	
	//서비스 목록
	interfaceMeta.setInterfaceServices(null != source? source.getInterfaceServices() : new ArrayList<>());
	
	// Property
	if(null != source) {
		interfaceMeta.setInterfaceProperties(source.getInterfaceProperties());
	} else {
		interfaceService.generateProperties(interfaceMeta, false);
	}
	
	//Response
	interfaceMeta.setInterfaceResponses(new ArrayList<>());
	
	return interfaceMeta;
	  
  }
  
  protected void generateDownload(HttpServletResponse response, String templateFile, InterfaceMeta entity, List<InterfaceMeta> list) throws Exception
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
    	    
      // ID
      values = entity.getId();
      row = writeSheet.getRow(3);
      cell = row.createCell(1);
      cell.setCellStyle(cellStyle_Base);
      cell.setCellValue(values);
    		
      // 이름
      values = entity.getName();
      cell = row.createCell(3);
      cell.setCellStyle(cellStyle_Base);
      cell.setCellValue(values);	
    		
      // Source 시스템 ID
      values = null != entity.getSourceSystemId() && 0 < entity.getSourceSystemId().trim().length()? entity.getSourceSystemId() : "";
      cell = row.createCell(5);
      cell.setCellStyle(cellStyle_Base);
      cell.setCellValue(values);
    		
      // 업무명
      values = entity.getBizId();
      cell = row.createCell(7);
      cell.setCellStyle(cellStyle_Base);
      cell.setCellValue(values);	

      //연계 서버 ID
      values = null != entity.getServerId() && 0 < entity.getServerId().trim().length()? entity.getServerId() : "";
      cell = row.createCell(9);
      cell.setCellStyle(cellStyle_Base);
      cell.setCellValue(values);
    		
      // 인터페이스 Class
      if (null != entity.getInterfaceClass() && 0 < entity.getInterfaceClass().trim().length()) {
    	  InterfaceClass interfaceClass = interfaceClassService.search(new InterfaceClass(), -1, null, false).stream()
											    	    .filter(interfaceClassInfo -> interfaceClassInfo.getId().equals(entity.getInterfaceClass()))
											    	    .collect(Collectors.toList())
											    	    .get(0);
    	  
    	  values = interfaceClass.getName();
      }

      cell = row.createCell(11);
      cell.setCellStyle(cellStyle_Base);
      cell.setCellValue(values);
    		
      // 서비스 ID
      values = entity.getServiceId();
      row = writeSheet.getRow(4);
      cell = row.createCell(1);
      cell.setCellStyle(cellStyle_Base);
      cell.setCellValue(values);
    		
      // 서비스 이름
      values = entity.getServiceName();
      cell = row.createCell(3);
      cell.setCellStyle(cellStyle_Base);
      cell.setCellValue(values);
    		
      // Target 시스템 ID
      values = null != entity.getTargetSystemId() && 0 < entity.getTargetSystemId().trim().length()? entity.getTargetSystemId() : "";
      cell = row.createCell(5);
      cell.setCellStyle(cellStyle_Base);
      cell.setCellValue(values);

      // 서비스 업무명
      values = entity.getServiceBizId();
      cell = row.createCell(7);
      cell.setCellStyle(cellStyle_Base);
      cell.setCellValue(values);
    		
      // 종류
      if(null != entity.getServiceInterfaceType() && 0 < entity.getServiceInterfaceType().trim().length())
        values = entity.getServiceInterfaceType();
      else 
    	values = "";
    		
      cell = row.createCell(9);
      cell.setCellStyle(cellStyle_Base);
      cell.setCellValue(values);	
    		
      // 사용 여부
      values = 'Y' == entity.getUseYn()? MessageGenerator.getMessage("label.yes", "yes") : 'N' == entity.getUseYn()? MessageGenerator.getMessage("label.no", "no") : "";
      cell = row.createCell(11);
      cell.setCellStyle(cellStyle_Base);
      cell.setCellValue(values);		
    		
      // 상태
      char publishStatus = entity.getPublishStatus();
      values = PublishModel.PUBLISH_MAKE == publishStatus? "MAKE" : 
	  	   	   PublishModel.PUBLISH_REQUEST == publishStatus? "REQUEST" : 
	  	   	   PublishModel.PUBLISH_CANCEL == publishStatus? "CANCEL" : 
	  	   	   PublishModel.PUBLISH_APPROVE == publishStatus? "APPROVE" : "";
      row = writeSheet.getRow(5);
      cell = row.createCell(1);
      cell.setCellStyle(cellStyle_Base);
      cell.setCellValue(values);		

      // 비고
      values = entity.getDescription();
      cell = row.createCell(3);
      cell.setCellStyle(cellStyle_Base);
      cell.setCellValue(values);	
    		
      // 조회리스트 입력
      long sum = 0;
      int i = 7;
    		
      for (InterfaceMeta interfaceMeta : list)
      {
        row = writeSheet.createRow(i);
    			  
        int c = 0;
    		  
        // ID
        values = interfaceMeta.getId();
        cell = row.createCell(c);
        cell.setCellStyle(cellStyle_Base);
        cell.setCellValue(values);
    		  
        // 이름
        values = interfaceMeta.getName();
        cell = row.createCell(++c);
        cell.setCellStyle(cellStyle_Base);
        cell.setCellValue(values);	
    		  
        // Source 시스템 ID
        values = interfaceMeta.getSourceSystemId();
        cell = row.createCell(++c);
        cell.setCellStyle(cellStyle_Base);
        cell.setCellValue(values);
    		  
        // 업무명
        values = interfaceMeta.getBizId();
        cell = row.createCell(++c);
        cell.setCellStyle(cellStyle_Base);
        cell.setCellValue(values);
    		  
        // 연계 서버 ID
        values = interfaceMeta.getServerId();
        cell = row.createCell(++c);
        cell.setCellStyle(cellStyle_Base);
        cell.setCellValue(values);
    		  
        // 상태
        publishStatus = interfaceMeta.getPublishStatus();
        values = PublishModel.PUBLISH_MAKE == publishStatus? "MAKE" : 
	  	   	     PublishModel.PUBLISH_REQUEST == publishStatus? "REQUEST" : 
	  	   	     PublishModel.PUBLISH_CANCEL == publishStatus? "CANCEL" : 
	  	   	     PublishModel.PUBLISH_APPROVE == publishStatus? "APPROVE" : "";
        cell = row.createCell(++c);
        cell.setCellStyle(cellStyle_Base);
        cell.setCellValue(values);	
        
        // 인터페이스 Class
        if (null != entity.getInterfaceClass() && 0 < entity.getInterfaceClass().trim().length()) {
      	  InterfaceClass interfaceClass = interfaceClassService.search(new InterfaceClass(), -1, null, false).stream()
							    	     .filter(interfaceClassInfo -> interfaceClassInfo.getId().equals(interfaceMeta.getInterfaceClass()))
							    	     .collect(Collectors.toList())
							    	     .get(0);
      	  
      	  values = interfaceClass.getName();
      	  
      	  cell = row.createCell(++c);
          cell.setCellStyle(cellStyle_Base);
          cell.setCellValue(values);
        }
    		  
        // 배포 요청 시간
        values = interfaceMeta.getPublishDateTime();
        cell = row.createCell(++c);
        cell.setCellStyle(cellStyle_Base);
        cell.setCellValue(values);	  
    		  
        // 배포 완료 시간
        values = null != interfaceMeta.getPublishTimestamp()? interfaceMeta.getPublishTimestamp().toString() : null;
        cell = row.createCell(++c);
        cell.setCellStyle(cellStyle_Base);
        cell.setCellValue(values);	
    		  
        // 사용 여부
        values = 'Y' == interfaceMeta.getUseYn()? MessageGenerator.getMessage("label.yes", "yes") : 'N' == interfaceMeta.getUseYn()? MessageGenerator.getMessage("label.no", "no") : MessageGenerator.getMessage("label.all", "All");
        cell = row.createCell(++c);
        cell.setCellStyle(cellStyle_Base);
        cell.setCellValue(values);
    		  
        // Version
        values = String.valueOf(interfaceMeta.getUpdateVersion());
        cell = row.createCell(++c);
        cell.setCellStyle(cellStyle_Base);
        cell.setCellValue(values);
    		  
        // 작성자
        values = interfaceMeta.getUpdateUserId();
        cell = row.createCell(++c);
        cell.setCellStyle(cellStyle_Base);
        cell.setCellValue(values);
    		  
        // 작성일
        values = null != interfaceMeta.getUpdateTimestamp()? interfaceMeta.getUpdateTimestamp().toString().substring(0, 19) : "";
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
  
  public String getIsEmptyRecordId(Workbook workbook, int sheetIdx) {
	  Sheet sheet = workbook.getSheetAt(sheetIdx);
	  
	  Row row = sheet.getRow(3);
	  Cell cell = row.getCell(1);

	  return cell.getStringCellValue();
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
    // 폰트
	Font font = workbook.createFont();
	font.setFontHeight((short) (20 * size));
	font.setFontName("굴림");
	font.setColor(color);

	return font;
  }  
}
