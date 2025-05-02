package link.wo.mysheetmcp.config;

import link.wo.mysheetmcp.service.Excel2JsonService;
import org.springframework.ai.tool.ToolCallbackProvider;
import org.springframework.ai.tool.method.MethodToolCallbackProvider;
import org.springframework.context.annotation.Bean;
import org.springframework.stereotype.Component;

@Component
public class ToolConfig {
    @Bean
    public ToolCallbackProvider myTools(Excel2JsonService weatherService) {
        return MethodToolCallbackProvider.builder()
                .toolObjects(weatherService)
                .build();
    }
}
