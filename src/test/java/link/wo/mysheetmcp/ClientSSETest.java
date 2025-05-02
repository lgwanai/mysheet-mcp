package link.wo.mysheetmcp;


import io.modelcontextprotocol.client.McpClient;
import io.modelcontextprotocol.client.transport.WebFluxSseClientTransport;
import io.modelcontextprotocol.spec.McpSchema;
import org.springframework.web.reactive.function.client.WebClient;

import java.io.File;
import java.util.Map;

public class ClientSSETest {
    public static void main(String[] args) {
        var transport = new WebFluxSseClientTransport(WebClient.builder().baseUrl("http://localhost:8080"));
        var client = McpClient.sync(transport).build();
        client.initialize();
        client.ping();
        // 列出并展示可用的工具
        McpSchema.ListToolsResult toolsList = client.listTools();
        System.out.println("可用工具 = " + toolsList);

        // 获取成都的天气
        McpSchema.CallToolResult weatherForecastResult = client.callTool(new McpSchema.CallToolRequest("excel2json",
                Map.of("excelFile", new File("/Users/wuliang/workspace/hotelXpress/src/test/resources/excelFile/2025年省专公司协议酒店报送信息表-昆明悦朗花园酒店.xls"))));
        System.out.println("返回结果: " + weatherForecastResult.content());

        client.closeGracefully();
    }
}
