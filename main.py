import smtplib
import os 
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
from unidecode import unidecode

#for replacing turkish words to english
turkish_to_english = {
    'ğ': 'g',
    'Ğ': 'G',
    'ş': 's',
    'Ş': 'S',
    'ı': 'i',
    'İ': 'I',
    'ç': 'c',
    'Ç': 'C',
    'ö': 'o',
    'Ö': 'O',
    'ü': 'u',
    'Ü': 'U',
    'â': 'a',
    'Â': 'A',
    'î': 'i',
    'Î': 'I',
    'û': 'u',
    'Û': 'U'
}

#replacing turkish words with english
def replace_turkish_chars_safe(s):
    if isinstance(s, str):
        for turkish_char, english_char in turkish_to_english.items():
            s = s.replace(turkish_char, english_char)
    return s


gonderilemeyenler = []
#reading the excel and extracting fullname and mails
file_path = 'C:/Users/Xigmatek-1/Desktop/mail/liste.xlsx'
df = pd.read_excel(file_path)
path_ser = 'C:/Users/Xigmatek-1/Desktop/mail/sertifika/'

isimsoyisim = (df['isim'] + " " + df['soyisim']).str.upper()
isimsoyisimenglish = isimsoyisim.apply(replace_turkish_chars_safe)
mail = (df['mail'])
#main loop
for i in range(377,399):
    
    try:
        #opening image
        with open(path_ser + isimsoyisimenglish[i] + ".png", 'rb') as f:
            img_data = f.read()

        #seting up the mail
        msg = MIMEMultipart()
        msg['Subject'] = 'Tretek Konferansı Katılımcı Belgesi'
        text = MIMEText(f"""Sayın {isimsoyisimenglish[i]},

        Tretek Konferansı'na katılımınız için teşekkür ederiz. Katılım belgenizin hazırlanmış olduğunu bildirmekten mutluluk duyuyoruz.

        Bu belge, Trakya Endüstri ve Teknoloji Konferansı'na katılımınızı resmi olarak doğrulayacacaktır.

        Etkinlik Detayları:

        Tarih: 16.12.23 17.12.23
        Saat: 10.00 17.00
        Konferans Yeri: Trakya Üniversitesi Devlet Konservatuar Salonu

        Katılımcı belgesi, sizin iki gün boyunca etkinliğimizdeki oturumların en az %80'ine katılım sağladığını temsil ediyor.

        Eğer herhangi bir düzeltme ya da değişiklik yapılması gerekiyorsa, lütfen bize en kısa sürede bildirin.

        Katılımınız için tekrar teşekkür ederiz. Eğer herhangi bir sorunuz ya da ihtiyacınız varsa, lütfen çekinmeden bizimle iletişime geçin.

        Saygılarımla,

        Talat Yılmaz
        IEEE Trakya Öğrenci Topluluğu RAS Komite Başkanı
        ieeetrakya@gmail.com
        ieeeturas@gmail.com 
        """)
        msg.attach(text)
        image = MIMEImage(img_data, name=os.path.basename(isimsoyisimenglish[i] + ".png"))
        msg.attach(image)


        #sending the mail
        server = smtplib.SMTP('smtp.gmail.com',587)
        server.starttls()
        server.login('mail','password')
        server.sendmail('mail', mail[i], msg.as_string()) 
        print(isimsoyisimenglish[i] + " Kişisine mail gönderildi.")
    except:
        print(isimsoyisimenglish[i] + "'in maili gönderilemedi.")
        gonderilemeyenler.append(isimsoyisimenglish[i])

#printing mail count to view how much mail sent
print(str(i + 1) + " tane mail başarıyla gönderildi.")
print(f"{gonderilemeyenler} Bu kisilere mail gonderilemedi")