package com;


import com.sun.net.httpserver.HttpServer;
import com.sun.net.httpserver.HttpHandler;
import com.sun.net.httpserver.HttpExchange;

import java.io.*;
import java.net.InetSocketAddress;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;

public class ExcelAutomationServer {

    private static final String TRACKER_PATH = "C:\\Users\\admin\\Desktop\\tracker\\Tracker.xlsx"; // Update this path
    private static HttpServer server;
    public static void main(String[] args) throws IOException, InterruptedException {
        server = HttpServer.create(new InetSocketAddress(5000), 0);
        server.createContext("/trigger", new TriggerHandler());
        server.setExecutor(null);
        server.start();
        System.out.println("✅ Server running at http://localhost:5000/trigger");
        HttpClient client = HttpClient.newHttpClient();
        HttpRequest request = HttpRequest.newBuilder()
                .uri(URI.create("http://localhost:5000/trigger"))
                .POST(HttpRequest.BodyPublishers.noBody())
                .build();
        client.send(request, HttpResponse.BodyHandlers.ofString());

    }

    static class TriggerHandler implements HttpHandler {
        @Override
        public void handle(HttpExchange exchange) throws IOException {
            if ("POST".equals(exchange.getRequestMethod())) {
                System.out.println("📨 Trigger received. Automating Excel...");
                automateExcel();
                String response = "Triggered";
                exchange.sendResponseHeaders(200, response.length());
                OutputStream os = exchange.getResponseBody();
                os.write(response.getBytes());
                os.close();

                System.out.println("🛑 Shutting down server after execution.");
                server.stop(0); // Gracefully stop the server

            }
        }

        private void automateExcel() {
            try {
                String scriptPath = "C:\\Users\\admin\\Desktop\\AutosaveExcelbatch\\automate_excel.ps1";
                ProcessBuilder pb = new ProcessBuilder("powershell.exe", "-ExecutionPolicy", "Bypass", "-File", scriptPath);
                pb.redirectErrorStream(true);
                Process process = pb.start();

                BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream()));
                String line;
                while ((line = reader.readLine()) != null) {
                    System.out.println("📄 PowerShell Output: " + line);
                }

                int exitCode = process.waitFor();
                System.out.println("🔚 PowerShell exited with code: " + exitCode);
            } catch (Exception e) {
                e.printStackTrace();
            }
        }

    }
}