function sendEmails() {
  // Aktif olan sayfayı alıyoruz
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Verileri alıyoruz (ilk satır başlık, veriler ise alt satırlarda)
  const data = sheet.getDataRange().getValues();
  
  // E-posta konusu
  const subject = "İyiliği Kodlayanlar Projesi Hakkında";
  
  // HTML formatında e-posta içeriği şablonu
  const emailBodyTemplate = `
    <p>Merhaba {{Name}},</p>

    <p>Öncelikle İyiliği Kodlayanlar projesine göstermiş olduğunuz ilgi ve değerli başvurunuz için teşekkür ederiz. Sizler gibi teknolojiye meraklı, öğrenmeye istekli bireylerin başvurularını incelemek bizim için çok kıymetliydi.</p>

    <p>Yapılan titiz değerlendirmeler sonucunda, ne yazık ki bu dönem projemize katılma şansı yakalayamadığınızı üzülerek paylaşmak istiyoruz.</p>

    <p>Ancak bu yolculuğun burada sona ermediğini unutmayın! Teknoloji dünyası sürekli öğrenim ve gelişimle dolu bir alan. Sizin gibi potansiyeli yüksek bireylerin gelecekte harika işlere imza atacağından eminiz. Projemiz kapsamında yeni dönem başvuruları ya da farklı fırsatlar için sizi bilgilendirmekten mutluluk duyarız. Bu minvalde bizi sosyal medya hesaplarımızdan takip etmeyi unutmayınız.</p>

    <p>Sizler de bu tarz projelerden haberdar olmak ya da içinde bulunmak istiyorsanız aşağıdaki iletişim ağı formunu doldurarak sürece dahil olabilirsiniz.</p>

    <p><a href="https://forms.gle/LnxwvpBepg7e8j3i9">https://forms.gle/LnxwvpBepg7e8j3i9</a></p>

    <p>İlginiz için tekrar teşekkür eder, gelecekte yollarımızın kesişmesini dileriz.</p>

    <p>Sevgilerimizle,<br>İyiliği Kodlayanlar Ekibi</p>
  `;
  
  // Her satır için e-posta gönderimi
  for (let i = 1; i < data.length; i++) {
    const name = data[i][0];  // İsim sütunu
    const email = data[i][3]; // E-posta sütunu
    
    // Şablondaki {{Name}} kısmını dinamik olarak isimle değiştiriyoruz
    const emailBody = emailBodyTemplate.replace("{{Name}}", name);
    
    try {
      // HTML formatında e-postayı gönderiyoruz
      GmailApp.sendEmail(email, subject, "", {
        htmlBody: emailBody
      });
      Logger.log("E-posta başarıyla gönderildi: " + email);  // Başarılı gönderimleri logla
    } catch (error) {
      Logger.log("E-posta gönderimi başarısız oldu: " + email + " - Hata: " + error.message);  // Hata oluşan e-posta adresini logla
    }
  }
}
