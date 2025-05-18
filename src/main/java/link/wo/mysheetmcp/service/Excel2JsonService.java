package link.wo.mysheetmcp.service;

import cn.hutool.core.io.FileUtil;
import cn.hutool.core.util.IdUtil;
import cn.hutool.core.util.StrUtil;
import cn.hutool.http.HttpUtil;
import cn.hutool.log.Log;
import cn.hutool.log.LogFactory;
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

    @Autowired
    Excel2JsonUtil excel2JsonUtil;

    private static final Log log = LogFactory.get();

    @Tool(description = "将excel文件转换成json")
    public String excel2Json(@ToolParam(description = "excel文件URL") String excelFileURL) {
        log.info("调用excel2Json方法,url:{}", excelFileURL);
        JSONObject json = new JSONObject();
        //excelFileURL是一个http开头的链接地址，需要下载到本地
        if (StrUtil.isEmpty(excelFileURL)) {
            return json.toJSONString();
        }
        // 获取文件名
        String fileName = excelFileURL.substring(excelFileURL.lastIndexOf('/') + 1);
        // 获取文件扩展名
        String fileExtension = fileName.substring(fileName.lastIndexOf('.') + 1);

        //生成一个以时间戳+四位随机数字的新文件名
        String newFileName = System.currentTimeMillis() + IdUtil.nanoId(4) + "." + fileExtension;
        // 将excelFile保存到STORAGE_FILE中
        File destFile = new File(STORAGE_FILE, newFileName);

        long size = HttpUtil.downloadFile(excelFileURL, FileUtil.file(destFile));

        if (size <= 0) {
            log.error("下载文件失败");
            return json.toJSONString();
        }

        try {
            json = excel2JsonUtil.toJson(destFile);
            log.info("excel2json:{}", json);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return json.toJSONString();
    }
}