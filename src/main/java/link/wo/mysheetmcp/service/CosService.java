package link.wo.mysheetmcp.service;

import com.qcloud.cos.COSClient;
import com.qcloud.cos.ClientConfig;
import com.qcloud.cos.auth.BasicCOSCredentials;
import com.qcloud.cos.auth.COSCredentials;
import com.qcloud.cos.model.PutObjectRequest;
import com.qcloud.cos.region.Region;
import jakarta.annotation.PostConstruct;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.InputStream;
import java.net.URL;
import java.util.Date;

@Service
public class CosService {

    @Value("${cos.secret-id}")
    private String secretId;

    @Value("${cos.secret-key}")
    private String secretKey;

    @Value("${cos.region}")
    private String regionName;

    @Value("${cos.bucket-name}")
    private String bucketName;

    private COSClient cosClient;

    @PostConstruct
    public void init() {
        COSCredentials cred = new BasicCOSCredentials(secretId, secretKey);
        Region region = new Region(regionName);
        ClientConfig clientConfig = new ClientConfig(region);
        cosClient = new COSClient(cred, clientConfig);
    }

    /**
     * 上传文件到 COS
     *
     * @param key      文件名 (包含路径)
     * @param localFile 本地文件
     * @return 文件的 URL
     */
    public String uploadFile(String key, File localFile) {
        PutObjectRequest putObjectRequest = new PutObjectRequest(bucketName, key, localFile);
        cosClient.putObject(putObjectRequest);
        return getUrl(key);
    }
    
    /**
     * 上传输入流到 COS
     * 
     * @param key 文件名 (包含路径)
     * @param inputStream 输入流
     * @param metadata 元数据 (可选)
     * @return 文件的 URL
     */
    public String uploadFile(String key, InputStream inputStream, com.qcloud.cos.model.ObjectMetadata metadata) {
        PutObjectRequest putObjectRequest = new PutObjectRequest(bucketName, key, inputStream, metadata);
        cosClient.putObject(putObjectRequest);
        return getUrl(key);
    }

    private String getUrl(String key) {
        // 使用 SDK 提供的方法获取 URL，这能更好地处理不同区域的域名差异
        URL url = cosClient.getObjectUrl(bucketName, key);
        return url.toString();
    }
    
    public void shutdown() {
        if (cosClient != null) {
            cosClient.shutdown();
        }
    }
}
