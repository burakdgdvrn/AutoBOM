# Havatek ERP BOM (Bill of Materials) Otomasyonu Proje Raporu

Bu rapor, projenin klasör yapısını, dosyaların görevlerini ve en önemlisi **Teknikerlerin PDF'ten veri çıkarırken gerçekte ne yaptığını** ile sistemimizin bunu nasıl dijitalleştirdiğini detaylıca açıklar.

---

## 1. Proje Klasör ve Dosya Yapısı

Proje temel olarak ikiye ayrılır: **Arka Plan (Backend/Python)** ve **Önyüz (Frontend/Web)**.

### Klasörler
*   📂 **`src/`**: Web arayüzünün (Önyüz) bulunduğu klasördür.
    *   `ERP Sistemi.html`: Kullanıcının gördüğü modern, karanlık/aydınlık (dark/light) temalı, profesyonel web arayüzünün iskeletidir.
    *   `custom_erp.css`: Web sitesinin o "SaaS (Kurumsal Yazılım)" premium hissiyatını veren stil ve renk tanımlamalarını içerir. Tablonun (Handsontable) renkleri ve buton tasarımları buradadır.
    *   `custom_erp.js`: Web sitesinin beynidir. PDF yüklenirken dönen çarkı göstermek, Python'dan gelen veriyi tabloya basmak, "Spul" (Assembly) atama açılır menüsünü (dropdown) yönetmek ve tabloyu kusursuz bir `.xlsx` dosyasına çevirmek (Excel'e Aktar) bu dosyanın görevidir.
*   📂 **`pdfler/`**: İşlenecek veya referans alınacak teknik çizimlerin (İzometrik PDF'lerin) bulunduğu klasördür.
*   📂 **`Excel/`**: Teknikerlerin ve mühendislerin kendi gözleriyle bakarak *manuel (elle)* oluşturdukları "Doğru (Ground Truth)" excel çıktılarını barındırır. Sistemin doğruluğunu test etmek için referans alınır.
*   📂 **`our_excel/`**: Bizim Python (Yapay zeka) sistemimizin PDF'leri okuyarak *otomatik* oluşturduğu excel çıktılarıdır. `Excel` klasörüyle kıyaslanarak hatalar tespit edilir.
*   📂 **`venv/`**: Python'un çalışması için gereken sanal kütüphane klasörüdür (Pytorch, EasyOCR, Flask vb. burada yüklüdür).

### Önemli Dosyalar
*   📄 **`app.py`**: Arayüz (HTML) ile Yapay Zeka (pdf_to_excel.py) arasındaki köprüdür (API Sunucusu). Arayüzden yüklenen PDF'i alır, bilgisayarın gizli geçici (temp) belleğine kaydeder (böylece VS Code sayfayı gereksiz yenilemez), ardından `pdf_to_excel` kodunu çalıştırıp sonucu arayüze geri yollar.
*   📄 **`pdf_to_excel.py`**: Projenin kalbidir. PDF'i bir resme çevirir, **EasyOCR** (Yapay Zeka) kullanarak resimdeki yazıları okur. Okunan dağınık yazıları koordinatlarına (`x` ve `y` düzleminde) bakarak satır ve sütunlara (Tabloya) oturtur ve ERP sisteminin anlayacağı formata (JSON) çevirir.
*   📄 **`Çalıştır.bat`**: Sistemi tek tıkla başlatan kısayol dosyasıdır.

---

## 2. İzometrik PDF'ten Neler, Neden ve Nasıl Çıkarılır?

Bir borulama (Piping) fabrikasında veya şantiyesinde, her şey İzometrik Çizimler (PDF) üzerinden yürür. Bu çizimlerin sağ tarafında veya köşesinde malzeme listeleri (BOM - Bill of Materials) bulunur.

### Hangi Tablolar Çekiliyor ve Neden?
1.  **CUT PIPE LENGTH (Kesik Boru Listesi):**
    *   **Neden:** Fabrikada uzun borular kesilerek çizime uygun hale getirilir. Ana borunun (Örn: Poz No 1) nasıl kesileceğini (Örn: 1-1, 1-2) ve hangi "Spool" (Montaj Parçası) numarasına ait olacağını belirtir.
2.  **FABRICATION MATERIALS (İmalat Malzemeleri):**
    *   **Neden:** Borular, dirsekler (Elbow), flanşlar (Flange) gibi atölye ortamında kaynatılıp birleştirilecek ana malzemelerin (Spool'ların) listesidir.
3.  **ERECTION MATERIALS (Montaj Malzemeleri):**
    *   **Neden:** Conta (Gasket), Cıvata (Bolt), Vana (Valve) gibi sahada (şantiyede) iki Spool'u birbirine bağlamak için gereken ek malzemelerin listesidir.

---

## 3. Teknikerlerin Gerçekte Yaptığı İşlem (Olması Gereken)

Teknikerler, ERP sistemine (veya Excel'e) veri girerken sadece tabloya körü körüne bakıp yazmazlar; **Mühendislik Yorumu** katarlar. İşte teknikerlerin elle yaptığı kusursuz işlemler:

1.  **Poz Numarası Eşleştirme (Sub Assembly):** Tablodaki `P/L NO` (Örn: 1, 2, 3) veya Cut Pipe tablosundaki `1-1, 1-2` numaralarını Excel'de `Sub Assembly` sütununa yazarlar.
2.  **Boruları (Pipes) Spool'lara Atama:** "Cut Pipe Length" tablosuna bakarak, örneğin "1-1 numaralı kesik boru SP01 spuluna aittir" bilgisini alıp `Assembly` sütununa `SP01` yazarlar.
3.  **Fittingsleri (Dirsek, Flanş) Spool'lara Atama (EN ZORU):** 
    *   **Durum:** PDF'teki tabloda hiçbir dirseğin veya flanşın hangi Spool'a ait olduğu *yazmaz*. 
    *   **Teknikerin Yaptığı:** Tekniker tabloya bakmayı bırakır, görsel çizimin içine dalar. Çizimdeki oklara, numaralara ve Weld (Kaynak) listesine bakar. "Hımm, 3 numaralı dirsek 1-1 numaralı boruya kaynatılmış. 1-1 borusu SP01 spulundaydı, demek ki 3 numaralı dirsek de SP01 spuluna (Assembly) ait olmalı" der ve Excel'e bunu elle işler.
4.  **Item Code Temizliği:** Çizimde "I1209668" veya sadece "1209668" yazsa da teknikerler şirketin ERP kültürüne göre bunu sadece "1209668" veya ERP sisteminde nasıl kayıtlıysa o şekilde standartlaştırıp girer.

---

## 4. Bizim Sistemimiz (Yapay Zeka) Bunu Nasıl Yapıyor?

Bizim yazdığımız Python (`pdf_to_excel.py`) scripti, teknikerin gözünün yaptığı bu "yorumlamayı" koda dökmeye çalışır:

1.  **Görseli Metne Çevirme:** PDF'i yüksek çözünürlüklü bir fotoğrafa çevirir ve EasyOCR ile kelimelerin `X` ve `Y` koordinatlarını bulur.
2.  **Sütunları Hizalama:** "Bu kelimenin X koordinatı %2 ile %18 arasındaysa bu kesin P/L NO sütunudur" diyerek havada uçuşan kelimeleri bir tabloya sıkıştırır.
3.  **Boruları (Pipes) Çözme:** Cut Pipe tablosunu okur. "1-1" borusunun Spool numarasını (Örn: SP01) hafızaya alır. Fabrication tablosunda o boruyu görünce Excel'e aktarırken `Assembly` kısmına `SP01` yazar. Tıpkı tekniker gibi!
4.  **Yapay Zekanın Sınırı ve Bizim Arayüz Çözümümüz (Fittings Spool Ataması):**
    *   Yapay zeka görsel çizimdeki çizgileri (topology) anlayamayacağı için dirsek ve flanşların Assembly (Spool) sütununu *otomatik olarak bulamaz.* (Çünkü tabloda yazmıyor).
    *   **Geliştirdiğimiz Çözüm:** Eğer PDF'te sadece tek bir Spool varsa, sistem "Bütün dirsekler mecburen buna aittir" deyip hepsini otomatik doldurur. Ancak çoklu Spool varsa boş bırakır. Burada topu teknikerin hızına atarız: Arayüzümüzde `Assembly` hücresine tıklandığında, sistemin hafızadaki spulları sunduğu bir **Açılır Menü (Dropdown)** çıkar. Tekniker saniyeler içinde "Bu SP01" deyip Excel'de olduğu gibi aşağı doğru kopyalayarak işlemi mükemmel şekilde tamamlar.
5.  **Otomatik Hata Düzeltme (Heuristic OCR):** Yapay zeka dikey tablo çizgisini `1` sanıp `1126` okursa, Python kodumuz teknikerin zihni gibi araya girer ve "Bu mantıksız, başındaki sahte 1'i I (İ) yap" der (`I126` veya standart neyse).

**Sonuç:** Bu sistem, dümdüz bir kopyala-yapıştır aracı değildir. Teknikerlerin "yorum" katarak yaptığı parçalama ve atama (BOM Data Entry) işlemlerinin algoritmik bir simülasyonudur. Sistemi ERP'ye dönüştüren asıl güç bu mantıktır.
