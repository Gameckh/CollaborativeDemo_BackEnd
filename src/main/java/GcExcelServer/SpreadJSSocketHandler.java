package GcExcelServer;

import com.google.gson.Gson;
import org.springframework.stereotype.Component;
import org.springframework.web.socket.BinaryMessage;
import org.springframework.web.socket.CloseStatus;
import org.springframework.web.socket.TextMessage;
import org.springframework.web.socket.WebSocketSession;
import org.springframework.web.socket.handler.AbstractWebSocketHandler;
import org.springframework.web.socket.handler.TextWebSocketHandler;

import java.io.IOException;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.CopyOnWriteArrayList;

@Component
public class SpreadJSSocketHandler extends AbstractWebSocketHandler {

    public SpreadJSSocketHandler(){

        // clean up the closed session every 5 minutes
        Timer timer = new Timer("cleanClosedSession");
        timer.schedule(new TimerTask() {
            @Override
            public void run() {
                Collection<List<WebSocketSession>> sessionsList = docSessions.values();
                for(List<WebSocketSession> sessions: sessionsList) {
                    ArrayList<WebSocketSession> closedSessions = new ArrayList<WebSocketSession>();
                    for(WebSocketSession session: sessions){
                        if(!session.isOpen()){
                            closedSessions.add(session);
                        }
                    }
                    sessions.removeAll(closedSessions);
                }
            }
        }, 5*60*1000,5*60*1000);
    }

    ConcurrentHashMap<String, List<WebSocketSession>> docSessions = new ConcurrentHashMap<String, List<WebSocketSession>>();

    @Override
    public void handleTextMessage(WebSocketSession session, TextMessage message)
        throws InterruptedException, IOException {
        Map value = new Gson().fromJson(message.getPayload(), Map.class);
        String cmd = (String) value.get("cmd");
        if(cmd == null){
            return;
        }
        String docID = (String) value.get("docID");
        if ("connect".equals(cmd.toLowerCase())) {
            List<WebSocketSession> sessions = null;
            if (docSessions.containsKey(docID)) {
                sessions = docSessions.get(docID);
            } else {
                sessions = Collections.synchronizedList(new ArrayList<WebSocketSession>());
                docSessions.put(docID, sessions);
            }
            sessions.add(session);
        } else {
            List<WebSocketSession> sessions = docSessions.get(docID);
            for (WebSocketSession webSocketSession : sessions) {
                if(webSocketSession.getId() != session.getId() &&
                    webSocketSession.isOpen()) {
                    webSocketSession.sendMessage(message);
                }
            }
        }
    }

//    @Override
//    protected void handleBinaryMessage(WebSocketSession session, BinaryMessage message) throws IOException {
//        System.out.println("New Binary Message Received");
//        session.sendMessage(message);
//    }

    @Override
    public void afterConnectionEstablished(WebSocketSession session) throws Exception {
        //the messages will be broadcasted to all users.
    }

    @Override
    public void afterConnectionClosed(WebSocketSession session, CloseStatus status) throws Exception {
//        session.close();
    }
}
