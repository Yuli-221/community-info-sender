import java.io.BufferedReader;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Properties;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import jakarta.mail.Authenticator;
import jakarta.mail.Message;
import jakarta.mail.PasswordAuthentication;
import jakarta.mail.Session;
import jakarta.mail.Transport;
import jakarta.mail.internet.InternetAddress;
import jakarta.mail.internet.MimeMessage;

public class ExcelReader {
  public static void main(String[] args) throws Exception {
    System.out.println(args[0]);
    System.out.println(args[1]);
    String mails = readEmailFromExcel();
    String mailContent = readResourceFile("mail-content.txt");

    System.out.println("===============");
    System.out.println("Mail 通知對象:");
    System.out.println(mails);

    System.out.println("===============");
    System.out.println("Mail 內容:");
    System.out.println(mailContent);
    System.out.println("===============");
    sendMail(args[0], args[1], "熱舞社-社團通知", mailContent, mails);

  }

  public static String readEmailFromExcel() throws Exception {
//    List<String> mails = new ArrayList<>();
    String mails = "";
    // 使用 ClassLoader 來取得 resources 中的 Excel 檔案
    ClassLoader classLoader = ExcelReader.class.getClassLoader();

    // 取得 resources 中的檔案路徑
    InputStream inputStream = classLoader.getResourceAsStream("迎新統計.xlsx");

    Sheet sheet = new XSSFWorkbook(inputStream).getSheetAt(0);
    Iterator<Row> it = sheet.iterator();
    // 跳過第一行
    it.next();

    while (it.hasNext()) {
      Row row = it.next();
      String mail = row.getCell(1).getStringCellValue();
      mails += mail+",";
    }
    return mails;
  }

  public static void sendMail(String mailAccount, String mailPws, String subject, String content, String mails)
      throws Exception {
    // 設定 SMTP 伺服器資訊
    String host = "smtp.gmail.com"; // Gmail SMTP 伺服器
    String port = "587"; // TLS 使用 587，SSL 使用 465
    String username = mailAccount; // 你的 Gmail 帳號
    String password = mailPws; // 應用程式專用密碼

    // 設定郵件屬性
    Properties props = new Properties();
    props.put("mail.smtp.auth", "true"); // 需要身份驗證
    props.put("mail.smtp.starttls.enable", "true"); // 啟用 TLS 加密
    props.put("mail.smtp.host", host);
    props.put("mail.smtp.port", port);

    // 建立 Session
    Session session = Session.getInstance(props, new Authenticator() {
      @Override
      protected PasswordAuthentication getPasswordAuthentication() {
        return new PasswordAuthentication(username, password);
      }
    });

    // 建立郵件內容
    Message message = new MimeMessage(session);
    message.setFrom(new InternetAddress(username)); // 寄件者
    message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(mails));
    message.setSubject(subject);
    message.setText(content);

    // 發送郵件
    Transport.send(message);
    System.out.println("郵件發送成功！");

  }

  public static String readResourceFile(String filePath) {
    ClassLoader classLoader = ExcelReader.class.getClassLoader();
    try (InputStream inputStream = classLoader.getResourceAsStream(filePath);
        BufferedReader reader = new BufferedReader(new InputStreamReader(inputStream, StandardCharsets.UTF_8))) {

      return reader.lines().collect(Collectors.joining("\n")); // 讀取所有行並合併成一個字串

    } catch (Exception e) {
      System.err.println("無法讀取檔案: " + filePath);
      e.printStackTrace();
    }
    return null;
  }
}
