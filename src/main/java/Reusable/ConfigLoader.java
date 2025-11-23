package Reusable;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;

public class ConfigLoader {
    private static Properties properties = new Properties();

    static {
        try {
            // Load from external file (same folder as JAR or a known path)
            String configPath = System.getProperty("config.file", "config.properties");
            FileInputStream fis = new FileInputStream(configPath);
            properties.load(fis);
        } catch (IOException e) {
            throw new RuntimeException("Failed to load config file", e);
        }
    }

    public static String get(String key) {
        return properties.getProperty(key);
    }
}