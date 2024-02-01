package org.example.model;

import javax.crypto.Cipher;
import javax.crypto.spec.SecretKeySpec;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathFactory;
import javax.xml.xpath.XPathExpression;

import org.example.libs.Response;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;

import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import java.io.File;
import java.sql.*;
import java.util.Base64;

import static org.example.model.DatabaseInitializer.createTablesIfNotExist;

public class DatabaseConnection {
    private String url;
    private String username;
    private String password;


    private int port;

    public String databaseType;
    String key = "beadc627d00ec777340bf6f06ece360fe1762e8b4408504516afd194dc303c77";
    public DatabaseConnection(String configFilePath) {
        try {
            // Parse the XML file
            DocumentBuilderFactory docBuilderFactory = DocumentBuilderFactory.newInstance();
            DocumentBuilder docBuilder = docBuilderFactory.newDocumentBuilder();
            Document doc = docBuilder.parse (new File(configFilePath));

            // normalize text representation
            doc.getDocumentElement().normalize();
            // Create XPath factory and XPath instance
            XPathFactory xpathFactory = XPathFactory.newInstance();
            XPath xpath = xpathFactory.newXPath();

            // Retrieve database configurations using XPath expressions
            databaseType = xpath.compile("/database/Database_Type").evaluate(doc);
            String databaseName = xpath.compile("/database/Database_Name").evaluate(doc);
            String host = xpath.compile("/database/Host").evaluate(doc);
              port = Integer.parseInt(xpath.compile("/database/Port").evaluate(doc));

            // Check if username and password are encrypted
            XPathExpression usernameExpr = xpath.compile("/database/Username/@encrypted");
            String usernameEncrypted = usernameExpr.evaluate(doc);

            XPathExpression passwordExpr = xpath.compile("/database/Password/@encrypted");
            String passwordEncrypted = passwordExpr.evaluate(doc);


            if ("yes".equals(usernameEncrypted)) {
                this.username = decrypt(xpath.compile("/database/Username").evaluate(doc), key);
            } else {
                this.username = xpath.compile("/database/Username").evaluate(doc);
                updateConfigFile(configFilePath, "/database/Username");
            }

            if ("yes".equals(passwordEncrypted)) {
                this.password = decrypt(xpath.compile("/database/Password").evaluate(doc), key);
            } else {
                this.password = xpath.compile("/database/Password").evaluate(doc);
                updateConfigFile(configFilePath, "/database/Password");
            }

            if ("Mysql".equalsIgnoreCase(databaseType)) {
                if (!checkDatabaseExists(host, port, databaseName)) {
                    createDatabase(host, port, databaseName);
                }
                this.url = "jdbc:mysql://" + host + ":" + port + "/" + databaseName;
            }else if("postgresql".equalsIgnoreCase(databaseType)){

                if (!checkDatabaseExists(host, port,databaseName, username, password)) {
                    createDatabasePostgreSQL(host,port, databaseName,username, password);
                }

                this.url = "jdbc:postgresql://" + host + ":" + port + "/" + databaseName+"?sslmode=require";
            }else if ("mssql".equalsIgnoreCase(databaseType)) {
                if (!checkDatabaseExistsMSSQL(host, port, databaseName, username, password)) {
                    createDatabaseMSSQL(host, port, databaseName, username, password);
                }
                this.url = "jdbc:sqlserver://" + host + ":" + port + ";databaseName=" + databaseName;
            }
            } catch (Exception e) {
             new Response(404,"Error: "+e.getMessage());
        }
             new Response(200,"Connected Successfully");
    }

    public Connection getConnection() throws Exception {
        Connection conn = DriverManager.getConnection(url, username, password);
       // createTablesIfNotExist(conn,databaseType);
        return  conn;
    }

    public String getDatabaseType(){
        return  this.databaseType;
    }
    private boolean checkDatabaseExistsMSSQL(String host, int port, String databaseName, String username, String password) {
        boolean exists = false;
        String url = "jdbc:sqlserver://" + host + ":" + port;
        try (Connection conn = DriverManager.getConnection(url, username, password);
             Statement stmt = conn.createStatement()) {

            String sql = "SELECT name FROM sys.databases WHERE name = '" + databaseName + "'";
            try (ResultSet rs = stmt.executeQuery(sql)) {
                exists = rs.next();
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return exists;
    }

    private void createDatabaseMSSQL(String host, int port, String databaseName, String username, String password) {
        String url = "jdbc:sqlserver://" + host + ":" + port;
        try (Connection conn = DriverManager.getConnection(url, username, password);
             Statement stmt = conn.createStatement()) {

            String sql = "CREATE DATABASE " + databaseName;
            stmt.executeUpdate(sql);
            System.out.println("Database created successfully");
        } catch (SQLException e) {
            if (e.getErrorCode() == 1801) {
                System.out.println("Database already exists");
            } else {
                e.printStackTrace();
            }
        }
    }
    private boolean checkDatabaseExists(String host, int port, String databaseName, String username, String password) {
        boolean exists = false;
        String url = "jdbc:postgresql://" + host + ":" + port + "/postgres?sslmode=require";

        try (Connection conn = DriverManager.getConnection(url, username, password);
             Statement stmt = conn.createStatement()) {

            String sql = "SELECT 1 FROM pg_database WHERE datname = '" + databaseName + "'";
            try (ResultSet rs = stmt.executeQuery(sql)) {
                exists = rs.next();
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return exists;
    }

    private void createDatabasePostgreSQL(String host, int port, String databaseName, String username, String password) {
        String url = "jdbc:postgresql://" + host + ":" + port + "/postgres?sslmode=require";

        try (Connection conn = DriverManager.getConnection(url, username, password);
             Statement stmt = conn.createStatement()) {

            String sql = "CREATE DATABASE \"" + databaseName + "\"";
            stmt.executeUpdate(sql);
            System.out.println("Database created successfully");
        } catch (SQLException e) {
            if (e.getSQLState().equals("42P04")) {
                System.out.println("Database already exists");
            } else {
                e.printStackTrace();
            }
        }
    }

    private boolean checkDatabaseExists(String host, int port, String databaseName) {
        try (Connection connection = DriverManager.getConnection("jdbc:mysql://" + host + ":" + port, username, password)) {
            ResultSet resultSet = connection.getMetaData().getCatalogs();
            while (resultSet.next()) {
                String dbName = resultSet.getString(1);
                if (dbName.equals(databaseName)) {
                    return true;
                }
            }
        } catch (Exception e) {
            new Response(500,"Error: "+e.getMessage());
        }
        return false;
    }
    private void createDatabase(String host, int port, String databaseName) {
        try (Connection connection = DriverManager.getConnection("jdbc:mysql://" + host + ":" + port, username, password);
             Statement statement = connection.createStatement()) {
            String createDatabaseSQL = "CREATE DATABASE " + databaseName;
            statement.executeUpdate(createDatabaseSQL);
            System.out.println("Database created successfully...");
        } catch (Exception e) {
            new Response(500,"Error: "+e.getMessage());
        }
    }

    public String encrypt(String data, String key) throws Exception {
        byte[] keyBytes = hexStringToByteArray(key);
        if (keyBytes.length != 32) {
            throw new IllegalArgumentException("Key must be 32 bytes long");
        }

        SecretKeySpec secretKey = new SecretKeySpec(keyBytes, "AES");
        Cipher cipher = Cipher.getInstance("AES");
        cipher.init(Cipher.ENCRYPT_MODE, secretKey);

        byte[] encrypted = cipher.doFinal(data.getBytes());
        return Base64.getEncoder().encodeToString(encrypted);
    }
    private String decrypt(String encryptedData, String key) throws Exception {
        byte[] keyBytes = hexStringToByteArray(key);
        if (keyBytes.length != 32) {
            throw new IllegalArgumentException("Key must be 32 bytes long");
        }

        SecretKeySpec secretKey = new SecretKeySpec(keyBytes, "AES");
        byte[] encryptedBytes = Base64.getDecoder().decode(encryptedData);
        Cipher cipher = Cipher.getInstance("AES");
        cipher.init(Cipher.DECRYPT_MODE, secretKey);

        byte[] decrypted = cipher.doFinal(encryptedBytes);

        return new String(decrypted);
    }
    public byte[] hexStringToByteArray(String s) {
        int len = s.length();
        byte[] data = new byte[len / 2];
        for (int i = 0; i < len; i += 2) {
            data[i / 2] = (byte) ((Character.digit(s.charAt(i), 16) << 4)
                    + Character.digit(s.charAt(i + 1), 16));
        }
        return data;
    }
    public Response updateConfigFile(String xml, String xpathExpression) {

        try {
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            DocumentBuilder builder = factory.newDocumentBuilder();
            Document document = builder.parse(xml);
            XPathFactory xpathFactory = XPathFactory.newInstance();
            XPath xpath = xpathFactory.newXPath();

            Node usernameNode = (Node) xpath.evaluate(xpathExpression, document, XPathConstants.NODE);
            if (usernameNode != null) {
                String username = usernameNode.getTextContent();
                String encryptedAttrUsername = ((Element) usernameNode).getAttribute("encrypted");
                if ("no".equalsIgnoreCase(encryptedAttrUsername)) {
                    username = encrypt(username,key);
                    ((Element) usernameNode).setTextContent(username);
                    ((Element) usernameNode).setAttribute("encrypted", "yes");
                }
            }

            Node passwordNode = (Node) xpath.evaluate(xpathExpression, document, XPathConstants.NODE);
            if (passwordNode != null) {
                String password = passwordNode.getTextContent();
                String encryptedAttrPassword = ((Element) passwordNode).getAttribute("encrypted");
                if ("no".equalsIgnoreCase(encryptedAttrPassword)) {
                    password = encrypt(password, key);
                    ((Element) passwordNode).setTextContent(password);
                    ((Element) passwordNode).setAttribute("encrypted", "yes");
                }
            }

            TransformerFactory transformerFactory = TransformerFactory.newInstance();
            Transformer transformer = transformerFactory.newTransformer();
            DOMSource source = new DOMSource(document);
            StreamResult result = new StreamResult(xml);
            transformer.transform(source, result);

        } catch (Exception e) {
            return new Response(404, "File error "+e.getMessage());
        }
        return new Response(200, "File Found");
    }

}
