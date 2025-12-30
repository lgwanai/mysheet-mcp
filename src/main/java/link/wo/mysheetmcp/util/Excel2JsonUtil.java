package link.wo.mysheetmcp.util;

import cn.hutool.log.Log;
import cn.hutool.log.LogFactory;
import com.alibaba.fastjson2.JSONArray;
import com.alibaba.fastjson2.JSONObject;
import link.wo.mysheetmcp.service.CosService;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import java.io.ByteArrayInputStream;
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
import java.util.List;
import java.util.ArrayList;
import java.util.concurrent.*;

@Component
public class Excel2JsonUtil {
    private static final Log log = LogFactory.get();
    private final DateTimeFormatter TIMESTAMP_FORMAT = DateTimeFormatter.ofPattern("yyyyMMddHHmmssSSS");
    @Value("${storage.file}")
    private String STORAGE_FILE_DIR;

    @Autowired
    private CosService cosService;

    public JSONObject toJson(File excelFile, String type) throws IOException {
        log.debug("调用 Excel2JsonUtil toJson()方法, type:{}", type);
        
        try (Workbook workbook = getWorkbook(excelFile)) {
            Map<String, String> fileMap = extractFilesFromExcel(workbook, excelFile.getName());
            log.info("fileMap:{}", fileMap);

            if ("row-object".equalsIgnoreCase(type)) {
                return toJsonRowObject(workbook, fileMap);
            }
            return toJsonBasic(workbook, fileMap);
        }
    }

    public JSONObject toJson(File excelFile) throws IOException {
        return toJson(excelFile, "basic");
    }

    private JSONObject toJsonBasic(Workbook workbook, Map<String, String> fileMap) {
        JSONObject json = new JSONObject();
        JSONArray data = new JSONArray();
        
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
                        // Use the COS URL directly
                        colObj.put("type", "file");
                        colObj.put("value", fileMap.get(colName+rowNum));
                        log.info("cell:{}  replaced ->  url:{}" , colName+rowNum, fileMap.get(colName+rowNum));
                    }else{
                        setCellValue(colObj, cell);
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
        json.put("data", data);
        return json;
    }

    private JSONObject toJsonRowObject(Workbook workbook, Map<String, String> fileMap) {
        JSONObject json = new JSONObject();
        JSONObject header = new JSONObject();
        JSONArray data = new JSONArray();

        if (workbook.getNumberOfSheets() > 0) {
            Sheet sheet = workbook.getSheetAt(0);
            int lastRowNum = sheet.getLastRowNum();

            // Process Header (Row 0)
            Row headerRow = sheet.getRow(0);
            int maxColIx = 0;
            if (headerRow != null) {
                maxColIx = headerRow.getLastCellNum();
                for (int i = 0; i < maxColIx; i++) {
                    Cell cell = headerRow.getCell(i);
                    String colName = getExcelColumnName(i) + "1";
                    if (cell != null) {
                        header.put(colName, cell.toString());
                    } else {
                        header.put(colName, "");
                    }
                }
            }

            // Process Data (Rows 1 to lastRowNum)
            for (int i = 1; i <= lastRowNum; i++) {
                Row row = sheet.getRow(i);
                JSONObject rowData = new JSONObject();
                rowData.put("index", i);

                for (int j = 0; j < maxColIx; j++) {
                    String key = getExcelColumnName(j) + "1";
                    JSONObject cellObj = new JSONObject();

                    int targetRow = i;
                    int targetCol = j;

                    if (sheet.getMergedRegions() != null) {
                        for (CellRangeAddress region : sheet.getMergedRegions()) {
                            if (region.isInRange(i, j)) {
                                targetRow = region.getFirstRow();
                                targetCol = region.getFirstColumn();
                                break;
                            }
                        }
                    }

                    Row srcRow = sheet.getRow(targetRow);
                    Cell srcCell = (srcRow != null) ? srcRow.getCell(targetCol) : null;
                    String srcColName = getExcelColumnName(targetCol);
                    String srcCoord = srcColName + (targetRow + 1);

                    if (!fileMap.isEmpty() && fileMap.containsKey(srcCoord)) {
                        cellObj.put("type", "file");
                        cellObj.put("value", fileMap.get(srcCoord));
                    } else {
                        setCellValue(cellObj, srcCell);
                    }

                    rowData.put(key, cellObj);
                }
                data.add(rowData);
            }
        }
        json.put("header", header);
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

    private void setCellValue(JSONObject colObj, Cell cell) {
        if (cell == null) {
            colObj.put("type", "text");
            colObj.put("value", "");
            return;
        }

        switch (cell.getCellType()) {
            case STRING -> {
                colObj.put("type", "text");
                colObj.put("value", cell.getStringCellValue());
            }
            case NUMERIC -> {
                if (DateUtil.isCellDateFormatted(cell)) {
                    colObj.put("type", "date");
                    // 使用 DataFormatter 格式化日期，或者手动格式化
                    // 用户要求 yyyy-mm-dd
                    java.util.Date date = cell.getDateCellValue();
                    if (date != null) {
                        java.text.SimpleDateFormat sdf = new java.text.SimpleDateFormat("yyyy-MM-dd");
                        colObj.put("value", sdf.format(date));
                    } else {
                        colObj.put("value", "");
                    }
                } else {
                    // Check for currency format
                    String formatString = cell.getCellStyle().getDataFormatString();
                    if (formatString != null && (formatString.contains("￥") || formatString.contains("$") || formatString.contains("€") || formatString.contains("£"))) {
                        colObj.put("type", "money");
                        DataFormatter dataFormatter = new DataFormatter();
                        colObj.put("value", dataFormatter.formatCellValue(cell));
                    } else {
                        colObj.put("type", "number");
                        colObj.put("value", cell.getNumericCellValue());
                    }
                }
            }
            case BOOLEAN -> {
                colObj.put("type", "boolean");
                colObj.put("value", cell.getBooleanCellValue());
            }
            case FORMULA -> {
                // 公式比较复杂，可能是数字、字符串等
                // 使用 CachedFormulaResultType
                switch (cell.getCachedFormulaResultType()) {
                    case STRING -> {
                        colObj.put("type", "text");
                        colObj.put("value", cell.getStringCellValue());
                    }
                    case NUMERIC -> {
                         // 公式里的日期判断比较麻烦，简单处理为数字
                         colObj.put("type", "number");
                         colObj.put("value", cell.getNumericCellValue());
                    }
                    case BOOLEAN -> {
                        colObj.put("type", "boolean");
                        colObj.put("value", cell.getBooleanCellValue());
                    }
                    default -> {
                        colObj.put("type", "text");
                        colObj.put("value", "");
                    }
                }
            }
            default -> {
                colObj.put("type", "text");
                colObj.put("value", "");
            }
        }
    }

    private Map<String, String> extractFilesFromExcel(Workbook workbook, String originalFileName) {
        Map<String, String> result = new ConcurrentHashMap<>();

        try (ExecutorService executor = Executors.newVirtualThreadPerTaskExecutor()) {
            List<CompletableFuture<Void>> futures = new ArrayList<>();

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
                                submitUploadTask(fileData, originalFileName, coord, executor, futures, result);
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
                                submitUploadTask(fileData, originalFileName, coord, executor, futures, result);
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
                                        submitUploadTask(fileData, originalFileName, coord, executor, futures, result);
                                    }
                                } else {
                                    log.debug("PackagePart is null or not available for XSSF object, trying fallback methods");
                                    // 尝试备用方法
                                    extractXSSFObjectDataFallback(obj, shape, originalFileName, result, executor, futures);
                                }
                            } catch (Exception e) {
                                log.error("Error extracting XSSF embedded file: {}", e.getMessage());
                                // 尝试备用方法
                                extractXSSFObjectDataFallback(obj, shape, originalFileName, result, executor, futures);
                            }
                        }// 处理HSSFPicture类型
                        else if (shape instanceof XSSFPicture picture) {
                            log.debug("Found HSSFPicture object");
                            try {
                                PictureData pictureData = picture.getPictureData();
                                byte[] fileData = pictureData.getData();
                                String coord = getShapeCoordinate(shape);
                                submitUploadTask(fileData, originalFileName, coord, executor, futures, result);
                            } catch (Exception e) {
                                log.error("Error extracting HSSFPicture: {}", e.getMessage(), e);
                            }
                        }
                    }
                }
            }

            if (!futures.isEmpty()) {
                CompletableFuture.allOf(futures.toArray(new CompletableFuture[0])).join();
            }
        }
        return result;
    }

    private void submitUploadTask(byte[] fileData, String originalFileName, String coord, ExecutorService executor, List<CompletableFuture<Void>> futures, Map<String, String> result) {
        CompletableFuture<Void> future = CompletableFuture.runAsync(() -> {
            try {
                String fileName = saveEmbeddedFile(fileData, originalFileName);
                result.put(coord, fileName);
                log.debug("Extracted embedded file at {} saved as {}", coord, fileName);
            } catch (IOException e) {
                log.error("Error uploading file for coord {}", coord, e);
            }
        }, executor);
        futures.add(future);
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

        // Upload to COS
        String key = "attachment/" + fileName;
        try (InputStream is = new ByteArrayInputStream(fileData)) {
            com.qcloud.cos.model.ObjectMetadata metadata = new com.qcloud.cos.model.ObjectMetadata();
            metadata.setContentLength(fileData.length);
            return cosService.uploadFile(key, is, metadata);
        }
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

    private void extractXSSFObjectDataFallback(XSSFObjectData obj, XSSFShape shape, String excelFileName, Map<String, String> result, ExecutorService executor, List<CompletableFuture<Void>> futures) {
        log.debug("Starting fallback extraction methods for XSSF object");
        try {
            // 尝试方法1：直接获取对象数据 - 最可靠的方法
            byte[] fileData = obj.getObjectData();
            if (fileData != null && fileData.length > 0) {
                String coord = getShapeCoordinate(shape);
                submitUploadTask(fileData, excelFileName, coord, executor, futures, result);
                log.debug("Extracted XSSF embedded file (fallback-1) at {} saved", coord);
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
                        submitUploadTask(fileData, excelFileName, coord, executor, futures, result);
                        log.debug("Extracted XSSF embedded file (fallback-2) at {} saved", coord);
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
                        submitUploadTask(fileData, excelFileName, coord, executor, futures, result);
                        log.debug("Extracted XSSF embedded file (fallback-3) at {} saved", coord);
                    }
                }
            }
        } catch (Exception e) {
            log.error("Fallback method 3 failed", e);
        }
    }
}
