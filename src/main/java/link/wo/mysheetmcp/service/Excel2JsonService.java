package link.wo.mysheetmcp.service;

import cn.hutool.cache.CacheUtil;
import cn.hutool.cache.impl.TimedCache;
import cn.hutool.crypto.digest.DigestUtil;
import cn.hutool.core.io.FileUtil;
import cn.hutool.core.util.IdUtil;
import cn.hutool.core.util.StrUtil;
import cn.hutool.http.HttpUtil;
import cn.hutool.log.Log;
import cn.hutool.log.LogFactory;
import com.alibaba.fastjson2.JSON;
import com.alibaba.fastjson2.JSONArray;
import com.alibaba.fastjson2.JSONObject;
import link.wo.mysheetmcp.util.Excel2JsonUtil;
import org.springframework.ai.tool.annotation.Tool;
import org.springframework.ai.tool.annotation.ToolParam;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.IOException;

@Service
public class Excel2JsonService {

    @Value("${storage.file}")
    private String STORAGE_FILE;
    @Value("${storage.cache}")
    private String STORAGE_CACHE;

    @Autowired
    Excel2JsonUtil excel2JsonUtil;

    private static final Log log = LogFactory.get();

    // Session Cache: Key = sessionId, Value = SessionData. Expire after 24 hours.
    private static final TimedCache<String, SessionData> sessionCache = CacheUtil.newTimedCache(24 * 60 * 60 * 1000);

    static {
        // Schedule pruning every hour to remove expired sessions
        sessionCache.schedulePrune(60 * 60 * 1000);
    }

    // Inner class to hold session data
    public static class SessionData {
        private String sessionId;
        private JSONArray dataRows;
        private JSONObject header;
        private int currentIndex;

        public SessionData(String sessionId, JSONArray dataRows, JSONObject header, int currentIndex) {
            this.sessionId = sessionId;
            this.dataRows = dataRows;
            this.header = header;
            this.currentIndex = currentIndex;
        }

        public String getSessionId() { return sessionId; }
        public JSONArray getDataRows() { return dataRows; }
        public JSONObject getHeader() { return header; }
        public int getCurrentIndex() { return currentIndex; }
        public void setCurrentIndex(int currentIndex) { this.currentIndex = currentIndex; }
    }

    @Tool(description = "Open an Excel file and create a read session. Returns a sessionId.")
    public JSONObject openFile(@ToolParam(description = "Excel file URL or local path") String url,
                           @ToolParam(description = "Reading mode: 'basic' or 'row-object'", required = false) String type,
                           @ToolParam(description = "Start reading from this line offset (default 0)", required = false) Integer offset) {
        log.info("Calling openFile, url:{}, type:{}, offset:{}", url, type, offset);
        
        // 1. Convert Excel to JSON
        JSONObject json = excel2Json(url, type);
        
        if (json == null || json.isEmpty()) {
            throw new RuntimeException("Failed to parse Excel file or file is empty.");
        }

        // 2. Prepare Data for Session
        JSONArray dataRows = new JSONArray();
        JSONObject header = null;

        if ("row-object".equalsIgnoreCase(type)) {
            if (json.containsKey("data")) {
                dataRows = json.getJSONArray("data");
            }
            if (json.containsKey("header")) {
                header = json.getJSONObject("header");
            }
        } else {
            // Default 'basic' mode: Flatten rows from all sheets
            if (json.containsKey("data")) {
                JSONArray sheets = json.getJSONArray("data");
                for (int i = 0; i < sheets.size(); i++) {
                    JSONObject sheet = sheets.getJSONObject(i);
                    if (sheet.containsKey("rows")) {
                        dataRows.addAll(sheet.getJSONArray("rows"));
                    }
                }
            }
        }

        // 3. Generate Session ID
        String sessionId = IdUtil.fastSimpleUUID();
        int startOffset = (offset != null && offset >= 0) ? offset : 0;

        // 4. Store in Cache
        SessionData sessionData = new SessionData(sessionId, dataRows, header, startOffset);
        sessionCache.put(sessionId, sessionData);

        log.info("Session created: {}, rows: {}", sessionId, dataRows.size());
        
        JSONObject result = new JSONObject();
        result.put("sessionId", sessionId);
        return result;
    }

    @Tool(description = "Read next line from the opened session. Returns header and current line content.")
    public JSONObject foreach(@ToolParam(description = "Session ID returned by openFile") String sessionId) {
        // 1. Get Session
        SessionData sessionData = sessionCache.get(sessionId);
        if (sessionData == null) {
            JSONObject error = new JSONObject();
            error.put("error", "Session expired or invalid");
            return error;
        }

        // 2. Refresh Expiry (by re-putting) - strictly speaking TimedCache get doesn't refresh, so we must put.
        // We will put it back at the end anyway after updating index.

        // 3. Read Current Line
        int index = sessionData.getCurrentIndex();
        JSONArray rows = sessionData.getDataRows();

        if (index >= rows.size()) {
            // End of file
            sessionCache.put(sessionId, sessionData); // Refresh expiry even if EOF? Yes, keep session alive.
            return new JSONObject(); // Return empty to indicate EOF
        }

        Object currentRow = rows.get(index);
        
        // 4. Update Index
        sessionData.setCurrentIndex(index + 1);
        sessionCache.put(sessionId, sessionData); // Update cache and refresh expiry

        // 5. Construct Response
        JSONObject result = new JSONObject();
        if (sessionData.getHeader() != null) {
            result.put("header", sessionData.getHeader());
        }
        result.put("row", currentRow);
        
        return result;
    }

    @Tool(description = "Reset the reading pointer to the beginning (0) for the given session.")
    public String reset(@ToolParam(description = "Session ID") String sessionId) {
        SessionData sessionData = sessionCache.get(sessionId);
        if (sessionData == null) {
            return "Error: Session expired or invalid";
        }
        
        sessionData.setCurrentIndex(0);
        sessionCache.put(sessionId, sessionData); // Update and refresh
        
        return "Success: Session reset to 0";
    }

    @Tool(description = "将excel文件转换成json")
    public JSONObject excel2Json(@ToolParam(description = "excel文件URL或本地路径") String excelFileURL,
                             @ToolParam(description = "解析模式：basic 或 row-object", required = false) String type) {
        log.info("调用excel2Json方法,url:{}, type:{}", excelFileURL, type);
        JSONObject json = new JSONObject();
        if (StrUtil.isEmpty(excelFileURL)) {
            return json;
        }

        File destFile;
        String fileName;

        if (excelFileURL.startsWith("http://") || excelFileURL.startsWith("https://")) {
            // 获取文件名
            fileName = excelFileURL.substring(excelFileURL.lastIndexOf('/') + 1);
            // 获取文件扩展名
            String fileExtension = fileName.substring(fileName.lastIndexOf('.') + 1);

            //生成一个以时间戳+四位随机数字的新文件名
            String newFileName = System.currentTimeMillis() + IdUtil.nanoId(4) + "." + fileExtension;
            // 将excelFile保存到STORAGE_FILE中
            destFile = new File(STORAGE_FILE, newFileName);
            // 确保目录存在
            if (!destFile.getParentFile().exists()) {
                destFile.getParentFile().mkdirs();
            }

            long size = HttpUtil.downloadFile(excelFileURL, FileUtil.file(destFile));

            if (size <= 0) {
                log.error("下载文件失败");
                return json;
            }
        } else {
            // 本地文件
            destFile = new File(excelFileURL);
            if (!destFile.exists()) {
                log.error("文件不存在: {}", excelFileURL);
                return json;
            }
            fileName = destFile.getName();
        }

        // Calculate MD5
        String md5 = DigestUtil.md5Hex(destFile);
        
        // Cache Logic
        File cacheDir = new File(STORAGE_CACHE);
        if (!cacheDir.exists()) {
            cacheDir.mkdirs();
        }
        String cacheFileName = md5 + (StrUtil.isEmpty(type) ? "" : "_" + type) + ".json";
        File cacheFile = new File(cacheDir, cacheFileName);

        if (cacheFile.exists()) {
            log.info("Cache hit for file: {}, md5: {}", fileName, md5);
            return JSON.parseObject(FileUtil.readUtf8String(cacheFile));
        }

        try {
            json = excel2JsonUtil.toJson(destFile, type);
            // Add metadata
            json.put("filename", fileName);
            json.put("md5", md5);
            
            // Save to cache
            FileUtil.writeUtf8String(json.toJSONString(), cacheFile);
            
            log.info("excel2json:{}", json);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return json;
    }
}