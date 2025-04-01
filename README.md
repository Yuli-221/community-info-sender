# 社團發送活動資訊應用程式

## 安裝 Java JDK 17
* 安裝 OpenJDK 17
## 安裝 Eclipse
* 安裝 Eclipse 

## 建立 Maven 專案
需要第三方函示庫進行 Mail 發送與 Excel 讀取，透過 Maven 有效使用第三方函示庫
* Apache POI : 讀取 Excel
* Jakarta Mail : 發送 Mail

Maven 載入函示庫宣告 pom.xml
```
	<dependencies>
		<dependency>
			<groupId>org.apache.poi</groupId>
			<artifactId>poi-ooxml</artifactId>
			<version>5.3.0</version>
		</dependency>
		<dependency>
			<groupId>com.sun.mail</groupId>
			<artifactId>jakarta.mail</artifactId>
			<version>2.0.1</version>
		</dependency>
	</dependencies>
```

## 研究 GMail 發送
* 因為預設 GMail 不允許 SMTP(Simple Mail Transfer Protocol, 簡單郵件傳輸協定) 發送需設定帳號安全
* 進入 Google 帳戶
* 選擇安全性 > 點選兩步驟驗證
![image](https://hackmd.io/_uploads/r1Uxd0uTJl.png)
![image](https://hackmd.io/_uploads/ryo7O0dTJe.png)
* 搜尋應用程式密碼並產生
![image](https://hackmd.io/_uploads/HkwRauF6kx.png)

* 發送 GMail 程式碼
* 採用 jakarta.mail.* 相關套件發送 Mail
* 因為要推送到 Github 上，所以進行密碼保護，在運行程式碼的時候才將帳號密碼帶入

程式碼
```java
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
```

## 研究讀取發送 Mail 的檔案
* 檔案存放在專案路徑內
* 採用相對路徑讀取檔案

程式碼
```java
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
```
* 檔案內容如下
```

韓樂X熱舞校内聯合迎新【所有"熱"愛都包"韓"你】

以下為行前通知大家都要注意一下窩!
【活動日期】
·112年11月4號星期六
【報到時間與地點】
·8:45-9:00/201(教大前棟一樓)
【攜帶物品】
·雨具
·健保卡
·環保餐具、水壺
·個人醫藥品
【其他事項】
·穿著壢中運動褲
．事先吃早餐才有力氣進行上午的行程
·最後!帶著一顆愉悅的心前來
有任何問題都歡迎私訊社帳
～所有熱愛都包韓你 一定要來哦！
```

## 讀取 Excel 檔案提取 Mail 資訊
* excel 檔案解析第一個 Sheet 中第二欄位

程式碼
```java
  public static List<String> readEmailFromExcel() throws Exception {
    List<String> mails = new ArrayList<>();
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
      mails.add(mail);
    }
    return mails;
  }
```

## 程式流程控制
* 為保護密碼，運行透過 args 在指令方式帶入
* 讀取 mail-content.txt Mail 內容
* 
```java
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
```

## 執行與結果 Console 內容

* 執行程式碼，帶入 Mail 帳號與應用程式密碼
![image](https://hackmd.io/_uploads/rJ3SoMY61g.png)

* 運行 Console 結果
```
Mail 通知對象:
s110005@student.clhs.tyc.edu.tw,
===============
Mail 內容:
韓樂X熱舞校内聯合迎新【所有"熱"愛都包"韓"你】

以下為行前通知大家都要注意一下窩!
【活動日期】
·112年11月4號星期六
【報到時間與地點】
·8:45-9:00/201(教大前棟一樓)
【攜帶物品】
·雨具
·健保卡
·環保餐具、水壺
·個人醫藥品
【其他事項】
·穿著壢中運動褲
．事先吃早餐才有力氣進行上午的行程
·最後!帶著一顆愉悅的心前來
有任何問題都歡迎私訊社帳
～所有熱愛都包韓你 一定要來哦！
===============
郵件發送成功！
```

## GitHub 推送程式碼
* 安裝 GitHub Desktop
  ![image](https://hackmd.io/_uploads/H1UTPdFayl.png)

* 登入 GitHub
  ![image](https://hackmd.io/_uploads/HyNB_uFaJg.png)

* 選擇以存在的本地專案
  ![image](https://hackmd.io/_uploads/Syw4oOt61g.png)

* 建立 GitHub 程式碼同步至遠端(Publish repository)
  ![image](https://hackmd.io/_uploads/H1P3i_F6yx.png)
![image](https://hackmd.io/_uploads/HJIRidF6Jg.png)
