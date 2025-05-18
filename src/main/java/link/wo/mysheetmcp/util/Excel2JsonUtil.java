package link.wo.mysheetmcp.util;

import cn.hutool.log.Log;
import cn.hutool.log.LogFactory;
import com.alibaba.fastjson2.JSONArray;
import com.alibaba.fastjson2.JSONObject;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.Map;

@Component
public class Excel2JsonUtil {
    private static final Log log = LogFactory.get();
    private final DateTimeFormatter TIMESTAMP_FORMAT = DateTimeFormatter.ofPattern("yyyyMMddHHmmssSSS");
    @Value("${storage.attachment}")
    private String STORAGE_ATTACHMENT_DIR;
    @Value("${storage.file}")
    private String STORAGE_FILE_DIR;
    @Value("${url.attachment}")
    private String URL_ATTACHMENT;
    public JSONObject toJson(File excelFile) throws IOException {
        log.debug("调用 Excel2JsonUtil toJson()方法");
        Map<String, String> fileMap = extractFilesFromExcel(excelFile);
        log.info("fileMap:{}",fileMap);
        JSONObject json = new JSONObject();
        JSONArray data = new JSONArray();
        try (Workbook workbook = getWorkbook(excelFile)) {
            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                Sheet sheet = workbook.getSheetAt(sheetIndex);
                JSONObject sheetObj = new JSONObject();
                sheetObj.put("sheet", sheet.getSheetName());
                JSONArray rows = new JSONArray();
                for (Row row : sheet) {
                    JSONObject rowObj = new JSONObject();
                    int rowNum = row.getRowNum()+1;
                    rowObj.put("rowIndex", rowNum);
                    JSONArray columns = new JSONArray();
                    for (Cell cell : row) {
                        JSONObject colObj = new JSONObject();
                        String colName = getExcelColumnName(cell.getColumnIndex());

                        colObj.put("colIndex", colName);
                        if(!fileMap.isEmpty() && fileMap.containsKey(colName+rowNum)){
                            colObj.put("value", "file::"+ URL_ATTACHMENT + fileMap.get(colName+rowNum));
                            log.info("value:{}  replaced ->  fileName:{}" , getCellValue(cell),fileMap.get(colName+rowNum));
                        }else{
                            colObj.put("value", getCellValue(cell));
                        }

                        if (sheet.getMergedRegions() != null) {
                            for (CellRangeAddress region : sheet.getMergedRegions()) {
                                if (region.isInRange(cell.getRowIndex(), cell.getColumnIndex())) {
                                    if (region.getLastRow() - region.getFirstRow() + 1 > 1) {
                                        colObj.put("rowspan", region.getLastRow() - region.getFirstRow() + 1);
                                    }
                                    if (region.getLastColumn() - region.getFirstColumn() + 1 > 1) {
                                        colObj.put("colspan", region.getLastColumn() - region.getFirstColumn() + 1);
                                    }
                                    break;
                                }
                            }
                        }
                        columns.add(colObj);
                    }
                    rowObj.put("columns", columns);
                    rows.add(rowObj);
                }
                sheetObj.put("rows", rows);
                data.add(sheetObj);
            }
        }
        json.put("data", data);
        return json;
    }

    private Workbook getWorkbook(File file) throws IOException {
        String fileName = file.getName();
        try (FileInputStream fis = new FileInputStream(file)) {
            if (fileName.endsWith(".xlsx")) {
                return new XSSFWorkbook(fis);
            } else if (fileName.endsWith(".xls")) {
                return new HSSFWorkbook(fis);
            } else {
                throw new IllegalArgumentException("Unsupported file format");
            }
        }
    }
    private String getExcelColumnName(int col) {
        StringBuilder columnName = new StringBuilder();
        while (col >= 0) {
            columnName.insert(0, (char)('A' + (col % 26)));
            col = (col / 26) - 1;
        }
        return columnName.toString();
    }

    private Object getCellValue(Cell cell) {
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> cell.getNumericCellValue();
            case BOOLEAN -> cell.getBooleanCellValue();
            case FORMULA -> cell.getCellFormula();
            default -> "";
        };
    }

    private Map<String, String> extractFilesFromExcel(File excelFile) throws IOException {
        Path tmpPath = Paths.get(STORAGE_ATTACHMENT_DIR);
        if (!Files.exists(tmpPath)) {
            Files.createDirectories(tmpPath);
        }

        Map<String, String> result = new HashMap<>();

        try (Workbook workbook = getWorkbook(excelFile)) {
            // 记录Excel文件类型信息
            log.debug("Excel file type: {}", workbook.getClass().getName());

            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                Sheet sheet = workbook.getSheetAt(sheetIndex);
                log.debug("Processing sheet: {}", sheet.getSheetName());

                // 获取绘图对象 - 注意：对于XLS文件，应该使用HSSFSheet的getDrawingPatriarch方法
                Drawing<?> drawing = null;
                if (workbook instanceof HSSFWorkbook) {
                    HSSFSheet hssfSheet = (HSSFSheet) sheet;
                    // 对于XLS文件，如果没有绘图对象，getDrawingPatriarch会创建一个
                    drawing = hssfSheet.getDrawingPatriarch();
                    log.debug("Using HSSFSheet.getDrawingPatriarch() for XLS file");
                } else {
                    drawing = sheet.getDrawingPatriarch();
                    log.debug("Using standard getDrawingPatriarch() for non-XLS file");
                }

                if (drawing == null) {
                    log.debug("No drawing found in sheet: {}", sheet.getSheetName());
                    continue;
                }

                log.debug("Drawing type: {}", drawing.getClass().getName());

                // 处理XLS格式
                if (drawing instanceof HSSFPatriarch patriarch) {
                    log.debug("Found HSSFPatriarch with {} children", patriarch.getChildren().size());
                    for (HSSFShape shape : patriarch.getChildren()) {
                        log.debug("Processing HSSFShape: {}", shape.getClass().getName());
                        // 先检查是否为HSSFObjectData类型
                        if (shape instanceof HSSFObjectData obj) {
                            log.debug("Found HSSFObjectData object");
                            try {
                                byte[] fileData = obj.getObjectData();
                                String coord = getShapeCoordinate(shape);
                                String fileName = saveEmbeddedFile(fileData, excelFile.getName());
                                result.put(coord, fileName);
                                log.debug("Extracted HSSF embedded file at {} saved as {}", coord, fileName);
                            } catch (Exception e) {
                                log.error("Error extracting HSSFObjectData: {}", e.getMessage(), e);
                            }
                        }
                        // 处理HSSFPicture类型
                        else if (shape instanceof HSSFPicture picture) {
                            log.debug("Found HSSFPicture object");
                            try {
                                PictureData pictureData = picture.getPictureData();
                                byte[] fileData = pictureData.getData();
                                String coord = getShapeCoordinate(shape);
                                String fileName = saveEmbeddedFile(fileData, excelFile.getName());
                                result.put(coord, fileName);
                                log.debug("Extracted HSSF picture at {} saved as {}", coord, fileName);
                            } catch (Exception e) {
                                log.error("Error extracting HSSFPicture: {}", e.getMessage(), e);
                            }
                        }
                        // 添加对其他类型的处理，确保不会因为类型不匹配而跳过后续逻辑
                        else {
                            log.debug("Skipping non-supported shape: {}", shape.getClass().getName());
                        }
                    }
                }
                // 处理XLSX格式
                else if (drawing instanceof XSSFDrawing xssfDrawing) {
                    for (XSSFShape shape : xssfDrawing.getShapes()) {
                        if (shape instanceof XSSFObjectData obj) {
                            try {
                                // 安全地获取PackagePart对象
                                PackagePart packagePart = null;
                                try {
                                    // 尝试使用反射安全地调用getPackagePart方法
                                    java.lang.reflect.Method getPackagePartMethod = obj.getClass().getMethod("getPackagePart");
                                    packagePart = (PackagePart) getPackagePartMethod.invoke(obj);
                                } catch (Exception e) {
                                    // 如果反射调用失败，记录日志但不抛出异常
                                    log.error("Reflection getPackagePart() call failed: {}", e.getMessage());
                                }

                                if (packagePart != null) {
                                    try (InputStream is = packagePart.getInputStream()) {
                                        byte[] fileData = is.readAllBytes();
                                        String coord = getShapeCoordinate(shape);
                                        String fileName = saveEmbeddedFile(fileData, excelFile.getName());
                                        result.put(coord, fileName);
                                        log.debug("Extracted XSSF embedded file at {} saved as {}", coord, fileName);
                                    }
                                } else {
                                    log.debug("PackagePart is null or not available for XSSF object, trying fallback methods");
                                    // 尝试备用方法
                                    extractXSSFObjectDataFallback(obj, shape, excelFile.getName(), result);
                                }
                            } catch (Exception e) {
                                log.error("Error extracting XSSF embedded file: {}", e.getMessage());
                                // 尝试备用方法
                                extractXSSFObjectDataFallback(obj, shape, excelFile.getName(), result);
                            }
                        }// 处理HSSFPicture类型
                        else if (shape instanceof XSSFPicture picture) {
                            log.debug("Found HSSFPicture object");
                            try {
                                PictureData pictureData = picture.getPictureData();
                                byte[] fileData = pictureData.getData();
                                String coord = getShapeCoordinate(shape);
                                String fileName = saveEmbeddedFile(fileData, excelFile.getName());
                                result.put(coord, fileName);
                                log.debug("Extracted HSSF picture at {} saved as {}", coord, fileName);
                            } catch (Exception e) {
                                log.error("Error extracting HSSFPicture: {}", e.getMessage(), e);
                            }
                        }
                    }
                }
            }
        }
        return result;
    }

    private String getShapeCoordinate(Object shape) {
        try {
            int row = 0;
            int col = 0;
            if (shape instanceof HSSFShape) {
                HSSFClientAnchor anchor = (HSSFClientAnchor) ((HSSFShape) shape).getAnchor();
                row = anchor.getRow1();
                col = anchor.getCol1();
            } else if (shape instanceof XSSFShape) {
                XSSFClientAnchor anchor = (XSSFClientAnchor) ((XSSFShape) shape).getAnchor();
                row = anchor.getRow1();
                col = anchor.getCol1();
            }
            return getExcelColumnName(col) + (row + 1);
        } catch (Exception e) {
            log.warn("Could not determine shape coordinate", e);
        }
        return "Unknown";
    }

    private String saveEmbeddedFile(byte[] fileData, String originalFileName) throws IOException {
        String timestamp = LocalDateTime.now().format(TIMESTAMP_FORMAT);
        String baseName = originalFileName.replaceAll("\\.[^.]+$", "");
        String extension = getFileExtension(fileData);
        String fileName = timestamp + "_" + baseName + extension;

        Path filePath = Paths.get(STORAGE_ATTACHMENT_DIR, fileName);
        Files.write(filePath, fileData);

        return fileName;
    }

    private String getFileExtension(byte[] fileData) {
        // 文件类型判断 - 通过文件头(Magic Numbers)识别
        if (fileData.length < 4) {
            return ".bin";
        }

        // PDF: %PDF (25 50 44 46)
        if (fileData[0] == (byte)0x25 && fileData[1] == (byte)0x50 &&
                fileData[2] == (byte)0x44 && fileData[3] == (byte)0x46) {
            return ".pdf";
        }
        // JPEG: FF D8 FF
        if (fileData[0] == (byte)0xFF && fileData[1] == (byte)0xD8 && fileData[2] == (byte)0xFF) {
            return ".jpg";
        }
        // PNG: 89 50 4E 47
        if (fileData[0] == (byte)0x89 && fileData[1] == (byte)0x50 &&
                fileData[2] == (byte)0x4E && fileData[3] == (byte)0x47) {
            return ".png";
        }
        // GIF: GIF8 (47 49 46 38)
        if (fileData[0] == (byte)0x47 && fileData[1] == (byte)0x49 &&
                fileData[2] == (byte)0x46 && fileData[3] == (byte)0x38) {
            return ".gif";
        }
        // DOC: D0 CF 11 E0
        if (fileData[0] == (byte)0xD0 && fileData[1] == (byte)0xCF &&
                fileData[2] == (byte)0x11 && fileData[3] == (byte)0xE0) {
            return ".doc";
        }
        // DOCX, XLSX, PPTX (ZIP format): PK
        if (fileData[0] == (byte)0x50 && fileData[1] == (byte)0x4B) {
            return ".docx"; // 默认为docx，实际可能是xlsx或pptx
        }
        // TXT: 检查是否为ASCII文本
        boolean isText = true;
        for (int i = 0; i < Math.min(fileData.length, 100); i++) {
            if (fileData[i] < 0x09 || (fileData[i] > 0x0D && fileData[i] < 0x20 && fileData[i] != 0x1B)) {
                isText = false;
                break;
            }
        }
        if (isText) {
            return ".txt";
        }

        return ".bin";
    }

    private void extractXSSFObjectDataFallback(XSSFObjectData obj, XSSFShape shape, String excelFileName, Map<String, String> result) {
        log.debug("Starting fallback extraction methods for XSSF object");
        try {
            // 尝试方法1：直接获取对象数据 - 最可靠的方法
            byte[] fileData = obj.getObjectData();
            if (fileData != null && fileData.length > 0) {
                String coord = getShapeCoordinate(shape);
                String fileName = saveEmbeddedFile(fileData, excelFileName);
                result.put(coord, fileName);
                log.debug("Extracted XSSF embedded file (fallback-1) at {} saved as {}", coord, fileName);
                return;
            } else {
                log.debug("Fallback method 1: Object data is null or empty");
            }
        } catch (Exception e) {
            log.error("Fallback method 1 failed: {}", e.getMessage());
        }

        try {
            // 尝试方法2：使用反射获取更多信息
            java.lang.reflect.Field field = obj.getClass().getDeclaredField("ctObject");
            field.setAccessible(true);
            Object ctObject = field.get(obj);

            if (ctObject != null) {
                // 尝试获取oleObject数据
                java.lang.reflect.Method getOleObject = ctObject.getClass().getMethod("getOleObject");
                Object oleObject = getOleObject.invoke(ctObject);

                if (oleObject != null) {
                    java.lang.reflect.Method getObjectData = oleObject.getClass().getMethod("getObjectData");
                    byte[] fileData = (byte[]) getObjectData.invoke(oleObject);

                    if (fileData != null && fileData.length > 0) {
                        String coord = getShapeCoordinate(shape);
                        String fileName = saveEmbeddedFile(fileData, excelFileName);
                        result.put(coord, fileName);
                        log.debug("Extracted XSSF embedded file (fallback-2) at {} saved as {}", coord, fileName);
                    }
                }
            }
        } catch (Exception e) {
            log.error("Fallback method 2 failed", e);
        }

        try {
            // 尝试方法3：处理OLE对象
            // 获取OLE对象的数据
            java.lang.reflect.Method getPreferredSize = obj.getClass().getMethod("getPreferredSize");
            Object preferredSize = getPreferredSize.invoke(obj);

            if (preferredSize != null) {
                // 尝试获取OLE对象的数据流
                java.lang.reflect.Field oleField = obj.getClass().getDeclaredField("ole");
                oleField.setAccessible(true);
                Object ole = oleField.get(obj);

                if (ole != null) {
                    java.lang.reflect.Method getBinaryData = ole.getClass().getMethod("getBinaryData");
                    byte[] fileData = (byte[]) getBinaryData.invoke(ole);

                    if (fileData != null && fileData.length > 0) {
                        String coord = getShapeCoordinate(shape);
                        String fileName = saveEmbeddedFile(fileData, excelFileName);
                        result.put(coord, fileName);
                        log.debug("Extracted XSSF embedded file (fallback-3) at {} saved as {}", coord, fileName);
                    }
                }
            }
        } catch (Exception e) {
            log.error("Fallback method 3 failed", e);
        }
    }
}
