package com.inzent.yellowpage.openapi.entity.record;

import java.io.FileInputStream;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.sql.Timestamp;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;
import java.util.Set;
import java.util.TreeSet;

import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;

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
import org.springframework.stereotype.Component;
import org.springframework.web.multipart.MultipartFile;

import com.fasterxml.jackson.core.JsonEncoding;
import com.inzent.igate.repository.meta.Field;
import com.inzent.igate.repository.meta.Record;
import com.inzent.imanager.message.MessageGenerator;
import com.inzent.yellowpage.controller.EntityExportImportBean;
import com.inzent.yellowpage.model.ModelField;
import com.inzent.yellowpage.model.ModelFieldPK;
import com.inzent.yellowpage.model.ModelRecord;
import com.inzent.yellowpage.model.PublishModel;

@Component
public class RecordExportImport implements EntityExportImportBean<ModelRecord> {
	
	@Autowired
	protected RecordRepository recordRepository;
	
	@Override
	public void exportList(HttpServletRequest request, HttpServletResponse response, ModelRecord entity, List<ModelRecord> list) throws Exception {
		String fileName = "RecordList_" + FastDateFormat.getInstance("yyyyMMdd_hhmmss").format(new Timestamp(System.currentTimeMillis())) + ".xlsx";

		response.addHeader("Cache-Control", "no-cache, no-store, must-revalidate");
		response.addHeader("Access-Control-Expose-Headers", "Content-Disposition");
		response.setHeader("Content-Disposition", "attachment; filename=\"" + fileName + "\"; filename*=UTF-8''" + URLEncoder.encode(fileName, JsonEncoding.UTF8.getJavaName()).replaceAll("\\+", "%20"));
		response.setContentType("application/octet-stream");

		generateDownload(response, request.getServletContext().getRealPath("/template/List_Record.xlsx"), entity, list);

		response.flushBuffer();
	}

	@Override
	public void exportObject(HttpServletRequest request, HttpServletResponse response, ModelRecord entity) throws Exception {
		String fileName = "Record_" + entity.getId() + "_" + FastDateFormat.getInstance("yyyyMMdd_hhmmss").format(new Timestamp(System.currentTimeMillis())) + ".xlsx";

		response.addHeader("Cache-Control", "no-cache, no-store, must-revalidate");
		response.addHeader("Access-Control-Expose-Headers", "Content-Disposition");
		response.setHeader("Content-Disposition", "attachment; filename=\"" + fileName + "\"; filename*=UTF-8''" + URLEncoder.encode(fileName, JsonEncoding.UTF8.getJavaName()).replaceAll("\\+", "%20"));
		response.setContentType("application/octet-stream");

		try (FileInputStream fileInputStream = new FileInputStream(request.getServletContext().getRealPath("/template/RecordTemplate.xlsx"));
			 Workbook workbook = WorkbookFactory.create(fileInputStream);
			 OutputStream outputStream = response.getOutputStream()) {
			
			exportExcelSheet(workbook, 0, entity);
			workbook.write(outputStream);
			
		} catch (Exception e) {
			throw e;
		}

		response.flushBuffer();
	}
	
	public void exportExcelSheet(Workbook workbook, int sheetIdx, ModelRecord entity)
	{
		Sheet writeSheet = workbook.getSheetAt(sheetIdx);
		Row row = null;
		Cell cell = null;

		CellStyle cellStyle_Base = getBaseCellStyle(workbook);

		row = writeSheet.getRow(3);

		// ID
		cell = row.getCell(1);
		cell.setCellValue(entity.getId());

		// 이름
		cell = row.getCell(4);
		cell.setCellValue(entity.getName());

		// 입&출력
		cell = row.getCell(9);
		String value = null;

		if (entity.getId().endsWith("_I")) 		 value = MessageGenerator.getMessage("label.input", "Input");
		else if (entity.getId().endsWith("_O"))  value = MessageGenerator.getMessage("label.output", "Output");
		else									 value = "";

		cell.setCellValue(value);

		// 권한
		cell = row.getCell(11);
		cell.setCellStyle(cellStyle_Base);
		cell.setCellValue(entity.getPrivilegeId());

		// Private
		cell = row.getCell(13);
		cell.setCellStyle(cellStyle_Base);
		cell.setCellValue(String.valueOf(entity.getPrivateYn()));
		
		// 종류
		cell = row.getCell(15);

		if (entity.getRecordType() == Record.TYPE_HEADER) 		value = MetaConstants.EXCEL_HEADER;
		else if (entity.getRecordType() == Record.TYPE_REFER)	value = MetaConstants.EXCEL_REFER;
		else													value = MetaConstants.EXCEL_INDIVI;

		cell.setCellValue(value);
		
		// MetaDomain
		cell = row.getCell(17);
		cell.setCellStyle(cellStyle_Base);
		cell.setCellValue(entity.getMetaDomain());
		
		row = writeSheet.getRow(4);
		
		// 사용여부
		cell = row.getCell(1);
		cell.setCellStyle(cellStyle_Base);
		cell.setCellValue(String.valueOf(entity.getUseYn()));

		// 옵션
		cell = row.getCell(3);
		cell.setCellStyle(cellStyle_Base);
		cell.setCellValue(entity.getRecordOptions());
		
		// 작성자
		cell = row.getCell(5);
		cell.setCellStyle(cellStyle_Base);
		cell.setCellValue(entity.getUpdateUserId());

		// 작성일
		cell = row.getCell(9);
		cell.setCellStyle(cellStyle_Base);
		cell.setCellValue(new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(null != entity.getUpdateTimestamp() ? entity.getUpdateTimestamp() : new Date(System.currentTimeMillis())));

		// 설명
		cell = row.getCell(13);
		cell.setCellStyle(cellStyle_Base);
		cell.setCellValue(entity.getDescription());

		exportExcelSheetRows(workbook, writeSheet, entity, 7, 0);
	}

	protected int exportExcelSheetRows(Workbook workbook, Sheet writeSheet, ModelRecord entity, int index, int depth) {
		Row row;
		Cell cell;

		// Cell 스타일 지정.
		CellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);// 텍스트 맞춤(세로 가운데)
		cellStyle.setAlignment(HorizontalAlignment.LEFT);// 텍스트 맞춤 (가로 왼쪽)

		Font font = workbook.createFont();// 폰트
		font.setFontHeight((short) 180);
		font.setFontName("굴림");
		font.setBold(false);
		cellStyle.setFont(font);

		cellStyle.setBorderBottom(BorderStyle.THIN);// Cell 테두리 (선)
		cellStyle.setBorderLeft(BorderStyle.THIN);
		cellStyle.setBorderRight(BorderStyle.THIN);
		cellStyle.setBorderTop(BorderStyle.THIN);

		cellStyle.setLocked(true);// Cell 잠금

		for (ModelField currentField : entity.getFields()) {
			
			// FIELD_LEVEL
			row = writeSheet.createRow(index);
			cell = row.createCell(0);
			cell.setCellStyle(cellStyle);
			cell.setCellValue(Integer.toString(depth));

			// 필드 ID
			String value = currentField.getPk().getFieldId();
			for (int j = 0; j < depth; j++)
				value = "   " + value;
			cell = row.createCell(1);
			cell.setCellStyle(cellStyle);
			cell.setCellValue(value);

			// FIELD_NAME
			cell = row.createCell(2);
			cell.setCellStyle(cellStyle);
			cell.setCellValue(currentField.getFieldName());

			// INDEX_FIELD_ID
			cell = row.createCell(3);
			cell.setCellStyle(cellStyle);
			cell.setCellValue(currentField.getFieldIndex());

			// FIELD_TYPE
			cell = row.createCell(4);
			cell.setCellStyle(cellStyle);
			cell.setCellValue(MetaConstants.FIELD_TYPES.get(currentField.getFieldType()));

			// FIELD_LENGTH
			cell = row.createCell(5);
			cell.setCellStyle(cellStyle);
			cell.setCellValue(Integer.toString(currentField.getFieldLength()));

			// FIELD_SCALE
			cell = row.createCell(6);
			cell.setCellStyle(cellStyle);
			cell.setCellValue(Integer.toString(currentField.getFieldScale()));

			// ARRAY_TYPE
			cell = row.createCell(7);
			cell.setCellStyle(cellStyle);
			cell.setCellValue(MetaConstants.FIELD_ARRAYTYPES.get(currentField.getArrayType()));

			// Reference_FIELD_ID
			cell = row.createCell(8);
			cell.setCellStyle(cellStyle);
			cell.setCellValue(currentField.getReferenceFieldId());

			// FIELD_DEFAULT_VALUE
			cell = row.createCell(9);
			cell.setCellStyle(cellStyle);
			cell.setCellValue(currentField.getFieldDefaultValue());

			// FIELD_HIDDEN_YN
			cell = row.createCell(10);
			cell.setCellStyle(cellStyle);
			cell.setCellValue(Character.toString(currentField.getFieldHiddenYn()));

			// FIELD_REQUIRE YN
			cell = row.createCell(11);
			cell.setCellStyle(cellStyle);
			cell.setCellValue(Character.toString(currentField.getFieldRequireYn()));

			// FIELD_VALID_VALUE
			cell = row.createCell(12);
			cell.setCellStyle(cellStyle);
			cell.setCellValue(currentField.getFieldValidValue());

			// FIELD_CODEC_ID
			cell = row.createCell(13);
			cell.setCellStyle(cellStyle);
			cell.setCellValue(currentField.getCodecId());

			// Options
			cell = row.createCell(14);
			cell.setCellStyle(cellStyle);
			cell.setCellValue(currentField.getFieldOptions());
			
			cell = row.createCell(15);
			cell.setCellStyle(cellStyle);
			
			cell = row.createCell(16);
			cell.setCellStyle(cellStyle);

			if (currentField.getFieldType() == Field.TYPE_RECORD && currentField.getSubRecordId() != null) {
				
				ModelRecord subRecord = currentField.getRecordObject();

				// RECORD_OPTION
				cell = row.createCell(14);
				cell.setCellStyle(cellStyle);
				cell.setCellValue(subRecord.getRecordOptions());

				index++;
				
				if (subRecord.getRecordType() == Record.TYPE_EMBED) {
					index = exportExcelSheetRows(workbook, writeSheet, subRecord, index, depth + 1);
				} else {
					
					cell = row.createCell(15);
					cell.setCellStyle(cellStyle);
					cell.setCellValue("Y");	

					// SUB_RECORD_ID
					cell = row.createCell(16);
					cell.setCellStyle(cellStyle);
					cell.setCellValue(currentField.getSubRecordId());
				}
			} else {
				index++;
			}
			
			// FIELD_DESC
			cell = row.createCell(17);
			cell.setCellStyle(cellStyle);
			cell.setCellValue(currentField.getFieldDesc());
		}

		return index;
	}

	@Override
	public ModelRecord importObject(MultipartFile multipartFile) throws Exception {
		
		ModelRecord modelRecord = new ModelRecord();

		try (OPCPackage opcPackage = OPCPackage.open(multipartFile.getInputStream()); Workbook workbook = new XSSFWorkbook(opcPackage)) {
			modelRecord = importExcelSheet(workbook, 0);			
		} catch (Exception e) {
			System.out.println(e);
			throw e;
		}

		return modelRecord;
	}
	
	public ModelRecord importExcelSheet(Workbook workbook, int sheetIdx) throws Exception
	{
		ModelRecord modelRecord = new ModelRecord();
		
		Sheet sheet = workbook.getSheetAt(sheetIdx);
		Row row = null;
	    Cell cell = null;
	    
		row = sheet.getRow(3);
		
		// ID
		cell = row.getCell(1);
		modelRecord.setId(cell.getStringCellValue());

		// 이름
		cell = row.getCell(4);
		modelRecord.setName(getStringNumericValue(cell));
		
		// 권한	
		cell = row.getCell(11);
		modelRecord.setPrivilegeId(getStringNumericValue(cell));
		
		// Private
		cell = row.getCell(13);
		modelRecord.setPrivateYn(getStringNumericValue(cell).charAt(0));

		// 모델유형
		cell = row.getCell(15);
		switch (getStringNumericValue(cell)) {
		case MetaConstants.EXCEL_HEADER:
			modelRecord.setRecordType(Record.TYPE_HEADER);
			break;

		case MetaConstants.EXCEL_REFER:
			modelRecord.setRecordType(Record.TYPE_REFER);
			break;

		default:
			modelRecord.setRecordType(Record.TYPE_INDIVI);
		}
		
		// MetaDomain
		cell = row.getCell(17);
		modelRecord.setMetaDomain(getStringNumericValue(cell));
		
		row = sheet.getRow(4);

		// 옵션
		cell = row.getCell(3);
		modelRecord.setRecordOptions(getStringNumericValue(cell));
		
		// 작성자
		cell = row.getCell(5);
		modelRecord.setUpdateUserId(getStringNumericValue(cell));
		
		// 작성일
		cell = row.getCell(9);
		modelRecord.setUpdateTimestamp(new Timestamp(new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").parse(getStringNumericValue(cell)).getTime()));
		
		// 설명
		cell = row.getCell(13);
		modelRecord.setDescription(getStringNumericValue(cell));
	
		// 배포 상태
		modelRecord.setPublishStatus(PublishModel.PUBLISH_MAKE);
		
		// 사용 여부
		modelRecord.setUseYn('Y');

		importExcelSheetRows(sheet, modelRecord, 7, 0);
	    
		return modelRecord;
	}

	protected int importExcelSheetRows(Sheet sheet, ModelRecord record, int index, int depth) throws Exception {
		Row row; 
		Cell cell;
		ModelField field;
		LinkedList<ModelField> filedList = new LinkedList<>();
		Set<String> duplicateIdList = new TreeSet<>();
		int nVal, idx = 0;

		while (true) {
			row = sheet.getRow(index);

			if (null == row || getStringNumericValue(row.getCell(1)).isEmpty()) break;

			// 필드가 있는 경우, Level 가져오기
			cell = row.getCell(0);
			nVal = Integer.parseInt(getStringNumericValue(cell));

			if (depth < nVal) // 윗줄 필드의 Level < 현재 필드의 Level
			{
				index = importExcelSheetRows(sheet, filedList.getLast().getRecordObject(), index, nVal);
				continue;
			} else if (depth > nVal) // 앞선 필드의 Level > 현재 필드의 Level
				break;

			field = new ModelField();
			field.setPk(new ModelFieldPK());
			field.getPk().setRecordId(record.getId());
			field.setFieldOrder(idx++) ;
			field.setRecord(record);

			// 필드 ID
			cell = row.getCell(1);
			field.getPk().setFieldId(getStringNumericValue(cell).trim());

			if (duplicateIdList.contains(field.getPk().getFieldId()))
				throw new Exception(MessageGenerator.getMessage("msg.duplicate.field.id", "Duplicate Field ID", field.getPk().getFieldId()));
			else
				duplicateIdList.add(field.getPk().getFieldId());

			// 필드 명
			cell = row.getCell(2);
			field.setFieldName(getStringNumericValue(cell));

			// 필드 Index
			cell = row.getCell(3);
			field.setFieldIndex(getStringNumericValue(cell));

			// 필드 타입 (오브젝트명)
			cell = row.getCell(4);
			field.setFieldType(MetaConstants.FIELD_TYPES_INVERT.get(getStringNumericValue(cell)));

			// 필드 길이
			cell = row.getCell(5);
			field.setFieldLength(Integer.parseInt(getStringNumericValue(cell)));

			// 필드 소수
			cell = row.getCell(6);
			field.setFieldScale(Integer.parseInt(getStringNumericValue(cell)));

			// 반복타입 (배열형태)
			cell = row.getCell(7);
			field.setArrayType(MetaConstants.FIELD_ARRAYTYPES_INVERT.get(getStringNumericValue(cell)));

			// 참조 필드 ID (반복횟수)
			cell = row.getCell(8);
			field.setReferenceFieldId(getStringNumericValue(cell));

			// 필드 기본값
			cell = row.getCell(9);
			field.setFieldDefaultValue(getStringNumericValue(cell));

			// 비공개여부 (마스킹여부)
			cell = row.getCell(10);
			field.setFieldHiddenYn(getStringNumericValue(cell).isEmpty()? 'N': getStringNumericValue(cell).charAt(0));

			// 필수여부
			cell = row.getCell(11);
			field.setFieldRequireYn(getStringNumericValue(cell).isEmpty()? 'N': getStringNumericValue(cell).charAt(0));

			// 유효값
			cell = row.getCell(12);
			field.setFieldValidValue(getStringNumericValue(cell));

			// 변환
			cell = row.getCell(13);
			field.setCodecId(getStringNumericValue(cell));

			// 기타속성
			cell = row.getCell(14);
			field.setFieldOptions(getStringNumericValue(cell));

			if (field.getFieldType() == Field.TYPE_RECORD) {
				ModelRecord subRecord = new ModelRecord();

				// 참조여부
				cell = row.getCell(15);
				if (getStringNumericValue(cell).equals("Y")) {
					String recordID = getStringNumericValue(row.getCell(16));
					subRecord.setId(recordID);

					try {
						subRecord = recordRepository.get(recordID);
					} catch (Exception e) {
						throw new Exception(MessageGenerator.getMessage("msg.unregister.refrence.model", "There is an unregistered reference model.", subRecord.getId()));
					}

					subRecord.setRecordType(Record.TYPE_REFER);
					field.setSubRecordId(recordID);
				} else {
					subRecord.setId(record.getId() + "@" + field.getPk().getFieldId());
					subRecord.setPrivilegeId(record.getPrivilegeId());
					subRecord.setRecordType(Record.TYPE_EMBED);
					
		            field.setSubRecordId(subRecord.getId());
				}

				field.setRecordObject(subRecord);

			}

			filedList.add(field);

			index++;
		}

		record.setFields(filedList);

		return index;

	}

	protected void generateDownload(HttpServletResponse response, String templateFile, ModelRecord entity,
			List<ModelRecord> list) throws Exception {
		try (FileInputStream fileInputStream = new FileInputStream(templateFile);
			Workbook workbook = WorkbookFactory.create(fileInputStream);
			OutputStream outputStream = response.getOutputStream();) {
			
			Sheet writeSheet = workbook.getSheetAt(0);
			Row row = null;
			Cell cell = null;
			String values = null;

			// Cell 스타일 지정.
			CellStyle cellStyle_Base = getBaseCellStyle(workbook);

			// 모델 ID
			values = entity.getId();
			row = writeSheet.getRow(3);
			cell = row.createCell(1);
			cell.setCellStyle(cellStyle_Base);
			cell.setCellValue(values);

			// 모델 이름
			values = entity.getName();
			row = writeSheet.getRow(3);
			cell = row.createCell(3);
			cell.setCellStyle(cellStyle_Base);
			cell.setCellValue(values);

			// 모델 종류
			char recordType = entity.getRecordType();
			values = 'H' == recordType ? "Header" : 'R' == recordType ? "Reference" : "";
			row = writeSheet.getRow(3);
			cell = row.createCell(5);
			cell.setCellStyle(cellStyle_Base);
			cell.setCellValue(values);

			// 사용 여부
			values = 'Y' == entity.getUseYn() ? MessageGenerator.getMessage("label.yes", "yes") : 'N' == entity.getUseYn() ? MessageGenerator.getMessage("label.no", "no") : "";
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

			// 비고
			values = entity.getDescription();
			row = writeSheet.getRow(4);
			cell = row.createCell(5);
			cell.setCellStyle(cellStyle_Base);
			cell.setCellValue(values);

			// 조회리스트 입력
			long sum = 0;
			int i = 6;

			for (ModelRecord modelRecord : list) {
				row = writeSheet.createRow(i);

				int c = 0;

				// 모델 ID
				values = modelRecord.getId();
				cell = row.createCell(c);
				cell.setCellStyle(cellStyle_Base);
				cell.setCellValue(values);

				// 모델 이름
				values = modelRecord.getName();
				cell = row.createCell(++c);
				cell.setCellStyle(cellStyle_Base);
				cell.setCellValue(values);

				// 모델 종류
				recordType = modelRecord.getRecordType();
				values = 'E' == recordType ? "Embed" : 'H' == recordType ? "Header" : 'I' == recordType ? "Individual" : 'R' == recordType ? "Reference" : 'C' == recordType ? "Common" : "";
				cell = row.createCell(++c);
				cell.setCellStyle(cellStyle_Base);
				cell.setCellValue(values);

				// 권한 ID
				values = modelRecord.getPrivilegeId();
				cell = row.createCell(++c);
				cell.setCellStyle(cellStyle_Base);
				cell.setCellValue(values);

				// 배포 상태
				publishStatus = modelRecord.getPublishStatus();
				values = PublishModel.PUBLISH_MAKE == publishStatus? "MAKE" : 
		    	  	     PublishModel.PUBLISH_REQUEST == publishStatus? "REQUEST" : 
		    	  	     PublishModel.PUBLISH_CANCEL == publishStatus? "CANCEL" : 
		    	  	     PublishModel.PUBLISH_APPROVE == publishStatus? "APPROVE" : "";
				cell = row.createCell(++c);
				cell.setCellStyle(cellStyle_Base);
				cell.setCellValue(values);

				// 배포 요청 시간
				values = modelRecord.getPublishDateTime();
				cell = row.createCell(++c);
				cell.setCellStyle(cellStyle_Base);
				cell.setCellValue(values);

				// 배포 완료 시간
				values = null != modelRecord.getPublishTimestamp() ? modelRecord.getPublishTimestamp().toString()
						: null;
				cell = row.createCell(++c);
				cell.setCellStyle(cellStyle_Base);
				cell.setCellValue(values);

				// 사용 여부
				values = 'Y' == modelRecord.getUseYn() ? MessageGenerator.getMessage("label.yes", "yes")
						: 'N' == modelRecord.getUseYn() ? MessageGenerator.getMessage("label.no", "no")
								: "";
				cell = row.createCell(++c);
				cell.setCellStyle(cellStyle_Base);
				cell.setCellValue(values);

				// Version
				values = String.valueOf(modelRecord.getUpdateVersion());
				cell = row.createCell(++c);
				cell.setCellStyle(cellStyle_Base);
				cell.setCellValue(values);

				// 작성자
				values = modelRecord.getUpdateUserId();
				cell = row.createCell(++c);
				cell.setCellStyle(cellStyle_Base);
				cell.setCellValue(values);

				// 작성일
				values = null != modelRecord.getUpdateTimestamp() ? modelRecord.getUpdateTimestamp().toString().substring(0, 19) : "";
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

			list = null;
			workbook.write(outputStream);
		} catch (Exception e) {
			throw e;
		}
	}

	public XSSFCellStyle getBaseCellStyle(Workbook workbook) {
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

	protected String getStringNumericValue(Cell cell) {
		if (cell != null)
			try {
				return cell.getStringCellValue();
			} catch (IllegalStateException e) {
				return Integer.toString((int) cell.getNumericCellValue());
			}

		return "";
	}

}