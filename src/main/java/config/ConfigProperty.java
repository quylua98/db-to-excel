package config;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

public class ConfigProperty {
    private ConfigProperty(){}
    private static final String CONFIG_FILE = "app.config";
    private static Properties properties = null;

    public static void init() throws IOException {
        Properties prop = new Properties();
        InputStream is = ConfigProperty.class.getClassLoader().getResourceAsStream(CONFIG_FILE);
        if(is != null) {
            prop.load(is);
        } else {
          throw new FileNotFoundException("");
        }
        properties = prop;
    }

    public static String getProperty(String key) throws IOException {
        if (properties == null) {
            init();
        }
        return (String) properties.get(key);
    }
}
