package link.wo.mysheetmcp.service;

import cn.hutool.crypto.digest.DigestUtil;
import cn.hutool.core.io.FileUtil;
import cn.hutool.core.util.IdUtil;
import cn.hutool.core.util.StrUtil;
import cn.hutool.http.HttpUtil;
import cn.hutool.log.Log;
import cn.hutool.log.LogFactory;
import com.alibaba.fastjson2.JSON;
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