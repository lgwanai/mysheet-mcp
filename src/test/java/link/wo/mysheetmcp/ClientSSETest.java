package link.wo.mysheetmcp;


import io.modelcontextprotocol.client.McpClient;
import io.modelcontextprotocol.client.transport.WebFluxSseClientTransport;
import io.modelcontextprotocol.spec.McpSchema;
import org.springframework.web.reactive.function.client.WebClient;

import java.io.File;
import java.util.Map;

public class ClientSSETest {
    public static void main(String[] args) {
//        var transport = new WebFluxSseClientTransport(WebClient.builder().baseUrl("http://mysheet.wo.link"));
        var transport = new WebFluxSseClientTransport(WebClient.builder().baseUrl("http://localhost:8080"));
        var client = McpClient.sync(transport).build();
        client.initialize();
        client.ping();
        // 列出并展示可用的工具
        McpSchema.ListToolsResult toolsList = client.listTools();
        System.out.println("可用工具 = " + toolsList);

        // 获取成都的天气
        McpSchema.CallToolResult weatherForecastResult = client.callTool(new McpSchema.CallToolRequest("excel2Json",
                Map.of("excelFileURL", "http://static.wo.link/upload/41f0ae83-7e43-4cb0-811d-22e51d820696.xlsx")));
        System.out.println("返回结果: " + weatherForecastResult.content());

        client.closeGracefully();
    }
}
