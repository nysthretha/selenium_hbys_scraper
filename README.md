# HBYS Sevk Takip Otomasyonu

Acil servisten sevk edilen hastaların, üst merkez (Uşak Eğitim ve Araştırma Hastanesi) tarafından yatışının yapılıp yapılmadığını otomatik olarak kontrol eden Selenium tabanlı bir otomasyon aracıdır.

## Gereksinimler

- **Windows 10/11** (Google Chrome yüklü)
- **Python 3.10+** — [python.org](https://www.python.org/downloads/) adresinden indirin, kurulumda **"Add to PATH"** kutucuğunu işaretleyin
- **ChromeDriver** — Chrome sürümünüze uygun olanı [buradan](https://googlechromelabs.github.io/chrome-for-testing/) indirin
- **Hastane ağı bağlantısı** — HBYS sistemi yalnızca hastane ağından erişilebilir

## Proje Yapısı

```
selenium_hbys_scraper/
├── scraper.py          # Ana otomasyon betiği
├── config.py           # Ayarlar: URL, sütun eşlemeleri, zamanlama
├── locators.py         # HBYS arayüzündeki buton/alan seçicileri (XPath, CSS)
├── .env                # Kullanıcı adı ve şifre (git'e eklenmez)
├── input.xlsx          # Girdi: sevk edilen hasta listesi
├── rapor.xlsx          # Çıktı: yatış sonuçları raporu
└── requirements.txt    # Python bağımlılıkları
```

### Dosya Açıklamaları

| Dosya | Açıklama |
|-------|----------|
| `scraper.py` | HBYS'ye giriş yapar, her hasta için TC ile arama yapar, Hasta Geçmişi penceresini açar, Tüm Hastaneler seçeneğini aktif eder ve yatış kaydı arar. Sonuçları `input.xlsx` dosyasına yazar ve `rapor.xlsx` raporunu oluşturur. |
| `config.py` | Tüm ayarları içerir: HBYS URL'si, Excel sütun eşlemeleri, bekleme süreleri, ChromeDriver yolu. |
| `locators.py` | HBYS arayüzündeki elementlerin seçicilerini (XPath/CSS) tanımlar. ExtJS dinamik ID'ler kullandığı için sabit class/name özelliklerine dayanır. |
| `.env` | HBYS kullanıcı adı ve şifresi. Bu dosya git'e dahil edilmez. |
| `input.xlsx` | Sevk edilen hastaların listesi. En az B sütununda TC Kimlik No ve G sütununda tarih olmalıdır. |
| `rapor.xlsx` | Otomasyon sonrası oluşturulan rapor. Yatışı olan hastalar yeşil, olmayanlar açık kahverengi ile renklenir. Sağ tarafta birim bazlı yatış sayıları yer alır. |

## Bilgisayarı Açtıktan Sonra Çalıştırma Adımları

### 1. ChromeDriver Kurulumu (yalnızca ilk seferde)

1. Chrome tarayıcınızın sürümünü öğrenin: Chrome > `⋮` > Yardım > Google Chrome Hakkında
2. Aynı sürüme uygun ChromeDriver'ı [indirin](https://googlechromelabs.github.io/chrome-for-testing/)
3. `chromedriver.exe` dosyasını `C:\hbys\chromedriver.exe` konumuna koyun

### 2. Python Ortamı Kurulumu (yalnızca ilk seferde)

PowerShell'i açın ve şu komutları çalıştırın:

```powershell
# Script çalıştırma izni (ilk seferde bir kez yeterli)
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# Sanal ortam oluştur
cd C:\hbys
python -m venv .venv
.venv\Scripts\activate

# Bağımlılıkları kur
pip install -r \\wsl$\Ubuntu\home\nysthretha\selenium_hbys_scraper\selenium_hbys_scraper\requirements.txt
```

### 3. `.env` Dosyasını Düzenleyin (şifre değiştiğinde)

Proje klasöründeki `.env` dosyasını açıp HBYS bilgilerinizi girin:

```
HBYS_USERNAME=KULLANICI_ADINIZ
HBYS_PASSWORD=SIFRENIZ
```

### 4. Girdi Dosyasını Hazırlayın

`input.xlsx` dosyasına sevk edilen hastaların listesini yerleştirin. Dosya şu sütunları içermelidir:

| Sütun | Alan |
|-------|------|
| A | ADI_SOYADI |
| B | TCKIMLIKNO |
| G | TARIH |
| H | TUMSEVKTANILARI |

### 5. Otomasyonu Çalıştırın

PowerShell'de:

```powershell
cd C:\hbys
.venv\Scripts\activate
python \\wsl$\Ubuntu\home\nysthretha\selenium_hbys_scraper\selenium_hbys_scraper\scraper.py
```

Chrome tarayıcısı otomatik açılır, HBYS'ye giriş yapar ve hastaları tek tek kontrol eder. **Tarayıcıya müdahale etmeyin.**

### 6. Raporu Görüntüleyin

Otomasyon tamamlandığında rapor dosyası şu konumda oluşur:

```
\\wsl$\Ubuntu\home\nysthretha\selenium_hbys_scraper\selenium_hbys_scraper\rapor.xlsx
```

Windows Gezgini adres çubuğuna yukarıdaki yolu yapıştırarak dosyayı açabilirsiniz.

## Rapor İçeriği

| Sütun | Açıklama |
|-------|----------|
| Tarih | Hastanın sevk edildiği tarih ve saat |
| Hasta Adı | Hastanın adı soyadı |
| Sevk Tanısı | Sevk sırasındaki tanı |
| Yatış Var | `Evet` / `Hayır` / `Hasta Bulunamadı` / `HATA` |
| Birim | Yatış yapılan birim (yalnızca Evet olanlar) |

Sağ tarafta (G-H sütunları) birim bazlı yatış özeti yer alır.

## Sık Karşılaşılan Sorunlar

| Sorun | Çözüm |
|-------|-------|
| `python` komutu bulunamıyor | Python kurulumunda "Add to PATH" işaretlenmemiş. Yeniden kurun veya `py` komutunu deneyin. |
| Chrome açılmıyor / chromedriver hatası | Chrome sürümünüz ile chromedriver sürümü uyuşmuyor. Güncel chromedriver indirin. |
| Sayfa yüklenmiyor / timeout | Hastane ağına bağlı olduğunuzdan emin olun. Ağ yavaşsa `config.py` içinde `WAIT_TIMEOUT` değerini artırın. |
| İlk hasta bulunamıyor | Sistem yavaş olabilir. `config.py` içinde `WAIT_TIMEOUT` değerini 60'a çıkarın. |
| Tüm sonuçlar HATA | Tarayıcı penceresi kapatılmış veya HBYS oturumu düşmüş olabilir. Tekrar çalıştırın. |
