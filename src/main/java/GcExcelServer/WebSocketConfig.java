package GcExcelServer;

import org.springframework.context.annotation.Configuration;
import org.springframework.web.socket.config.annotation.EnableWebSocket;
import org.springframework.web.socket.config.annotation.WebSocketConfigurer;
import org.springframework.web.socket.config.annotation.WebSocketHandlerRegistry;

import java.util.logging.SocketHandler;

@Configuration
@EnableWebSocket
public class WebSocketConfig implements WebSocketConfigurer {
    private SpreadJSSocketHandler handler = new SpreadJSSocketHandler();

    public void registerWebSocketHandlers(WebSocketHandlerRegistry registry) {
        registry.addHandler(this.handler, "/spreadjs").setAllowedOrigins("*");
    }

    public SpreadJSSocketHandler getSpreadJSSocketHandler(){
        return this.handler;
    }
}
