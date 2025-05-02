package link.wo.mysheetmcp.service;

import cn.hutool.core.io.FileUtil;
import cn.hutool.core.util.IdUtil;
import cn.hutool.core.util.RandomUtil;
import cn.hutool.log.Log;
import cn.hutool.log.LogFactory;
import com.alibaba.fastjson2.JSONObject;
import link.wo.mysheetmcp.util.Excel2JsonUtil;
import org.apache.xmlbeans.impl.common.IOUtil;
import org.springframework.ai.tool.annotation.Tool;
import org.springframework.ai.tool.annotation.ToolParam;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.IOException;

@Service
public class Excel2JsonService {
    @Value("${upload.tmp}")
    private String TMP_DIR;
    @Value("${upload.path}")
    private String UPLOAD_DIR;

    @Autowired
    Excel2JsonUtil excel2JsonUtil;

    private static final Log log = LogFactory.get();

    @Tool(description = "将excel文件转换成json")
    public String excel2Json(@ToolParam(description = "excel文件") File excelFile) {
        log.info("调用excel2Json方法");
        //取excelFile的后缀名
        String suffix = FileUtil.getSuffix(excelFile);
        //生成一个以时间戳+四位随机数字的新文件名
        String newFileName = System.currentTimeMillis() + IdUtil.nanoId(4)+"."+suffix;
        // 将excelFile保存到UPLOAD_DIR中
        File destFile = new File(UPLOAD_DIR, newFileName);
        FileUtil.copy(excelFile, destFile, true);
        log.info("保存文件到{}", destFile.getAbsolutePath());
        JSONObject json = new JSONObject();
        try {
            json = excel2JsonUtil.toJson(excelFile);
            log.info("excel2json:{}", json);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return json.toJSONString();
    }
}