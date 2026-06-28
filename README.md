# 🎓 Online Kurs Platformu

Kurs yönetimi, öğrenci takibi, eğitmen-mentor ilişkileri ve gamification içeren kapsamlı bir online eğitim platformu. PyQt5 ile geliştirilmiş, SQLite veritabanı destekli, Admin ve Öğrenci olmak üzere iki ayrı arayüze sahiptir.

**🔐 Giriş Bilgileri:**
- Admin: `admin` / `admin123`
- Öğrenci: Örnek öğrenci hesapları otomatik oluşturulur (sistemden görüntülenebilir)

---

## 📸 Ekran Görüntüleri

<img width="258" height="315" alt="image" src="https://github.com/user-attachments/assets/0f7d6aeb-521e-481e-a2e8-ec496f655bc7" />


---

## 📋 Özellikler

---

### 📊 Dashboard (Admin)

- Toplam kurs, öğrenci, eğitmen sayısı KPI kartları
- Aktif kayıt ve tamamlanma oranı
- QPainter bar grafik (kurs doluluk durumu)
- QPainter pasta grafik (kayıt durum dağılımı)
- Son aktiviteler

<img width="1279" height="757" alt="image" src="https://github.com/user-attachments/assets/b0a41578-b564-4977-b7d5-0136731c0fae" />


---

### 📚 Kurslar

- Kurs ekleme, düzenleme, silme
- Kurs adı, eğitmen, kategori, kontenjan, fiyat, açıklama, süre
- Kontenjan doluluk takibi
- Anlık arama ve kategori filtresi

<img width="1279" height="748" alt="image" src="https://github.com/user-attachments/assets/2c69cec3-6e74-4b47-8cba-64186ff14bd5" />


---

### 👥 Öğrenciler

- Öğrenci listesi ve detay görünümü
- Ad, soyad, email, kayıt tarihi, paket, XP, level
- Öğrenciye kayıtlı kursları görüntüleme
- Anlık arama

<img width="1279" height="751" alt="image" src="https://github.com/user-attachments/assets/f621e625-af9a-421d-8a4b-ac7240f7a302" />


---

### 👨‍🏫 Eğitmenler

- Eğitmen ekleme, düzenleme
- Ad, uzmanlık, email, biyografi, rating
- Eğitmene ait kurs listesi
- Ortalama rating gösterimi

<img width="707" height="434" alt="image" src="https://github.com/user-attachments/assets/0bd0613c-4105-4404-ab21-a48ae27e2bf8" />


---

### 🎮 Gamification

- XP (deneyim puanı) sistemi — kurs tamamlama, görev, rozet ile kazanım
- Level sistemi (XP eşiklerine göre otomatik level atlama)
- Rozet (Badge) sistemi — kriterlere göre otomatik kazanma
- XP geçmişi listesi
- İlerleme çubuğu (bir sonraki level için)

<img width="691" height="428" alt="image" src="https://github.com/user-attachments/assets/e221c8d1-1bfb-4cf7-a183-d2b072777eff" />


---

### 🧑‍🏫 Mentor & Sertifika

- Eğitmen-öğrenci mentor ilişkisi tanımlama
- Mentor görüşme tarihi ve notları
- Kurs tamamlayan öğrencilere sertifika oluşturma
- Sertifika durumu takibi


<img width="1279" height="296" alt="image" src="https://github.com/user-attachments/assets/fc7d8ecf-bdcc-4fb1-9655-5225e18aa496" />



<img width="1279" height="299" alt="image" src="https://github.com/user-attachments/assets/588dd857-053a-4c53-a4f7-8a83000fb62b" />

---

### 📈 Grafikler

- Kurs bazlı kayıt sayısı bar grafiği (QPainter)
- Aylık kayıt trendi çizgi grafiği (QPainter)
- Öğrenci seviye dağılımı pasta grafiği (QPainter)

<img width="692" height="454" alt="image" src="https://github.com/user-attachments/assets/bda37aab-c0fe-414b-bf8d-3e668e5f9ea9" />


<img width="701" height="454" alt="image" src="https://github.com/user-attachments/assets/d62733ea-659f-4847-ade5-61e84f51f46f" />



<img width="701" height="448" alt="image" src="https://github.com/user-attachments/assets/c511ef96-b6e4-49bf-8d23-5dc35d663afe" />



---

### 📋 Raporlar

- Kurs doluluk raporu
- Öğrenci ilerleme raporu
- Eğitmen performans raporu
- CSV dışa aktarma

<img width="1278" height="658" alt="image" src="https://github.com/user-attachments/assets/0d1f77f8-41bf-4534-a045-e36ed5db83eb" />


---

### ⚙️ Ayarlar

- Platform adı, iletişim bilgileri
- XP ve level eşik değerleri
- Sertifika şablonu metni
- DB yedekleme

<img width="1279" height="487" alt="image" src="https://github.com/user-attachments/assets/c291b563-1ce0-4af6-b5ab-311276a781b2" />


---

### 🎓 Öğrenci Görünümü

- Öğrenciye özel arayüz (farklı sekme düzeni)
- Kayıtlı kursları ve ilerleme yüzdesi
- XP, level ve rozet görünümü
- Mentor bilgisi
- Sertifikaları

<img width="1272" height="640" alt="image" src="https://github.com/user-attachments/assets/858684c1-b777-4665-aefc-88f8a8b9a199" />




<img width="1279" height="607" alt="image" src="https://github.com/user-attachments/assets/58b46293-082f-48ca-8dd3-a21cbe612a81" />







---

## ❓ Final Soruları

### Sistemde hangi kullanıcılar veya nesneler vardır?

**Kullanıcı Türleri:**
- **Admin** — Tam yetkili; kurs, öğrenci, eğitmen yönetimi, raporlar, ayarlar
- **Öğrenci** — Kayıtlı kurslara erişim, ilerleme takibi, XP/rozet görünümü

**Ana Nesneler / Varlıklar:**
- **Kurs** — kurs_id, ad, eğitmen, kategori, kontenjan, fiyat, açıklama, süre, durum
- **Öğrenci** — ogrenci_id, ad, soyad, email, kullanıcı adı, paket, XP, level, kayıt tarihi
- **Eğitmen** — egitmen_id, ad, uzmanlık, email, biyografi, rating, durum
- **Kayıt** — kayit_id, öğrenci, kurs, durum, ilerleme, XP kazanılan, kayıt tarihi
- **Enhancement Kayıt** — öğrenci, kurs, enhancement kodu, fiyat
- **XP Hareketi** — ogrenci_id, işlem türü, XP miktarı, açıklama, tarih
- **Badge (Rozet)** — badge_id, kod, ad, açıklama, ikon, XP değeri, kriter
- **Badge Kaydı** — öğrenci, badge, kazanma tarihi
- **Sertifika** — sertifika_id, öğrenci, kurs, oluşturma tarihi, durum
- **Mentor İlişkisi** — mentor_id, eğitmen, öğrenci, başlangıç tarihi, notlar
- **Aktivite Logu** — ogrenci_id, işlem türü, tarih
- **Sistem Kullanıcısı** — kullanici_id, ad, soyad, kullanıcı adı, SHA256 parola, rol

---

### Kullanıcı sistemde hangi işlemleri gerçekleştirebilir?

**Admin:**
- Kurs ekleyebilir, düzenleyebilir, silebilir
- Öğrenci listesini görüntüleyebilir, arayabilir, detay inceleyebilir
- Eğitmen ekleyebilir, düzenleyebilir
- Gamification ayarlarını (XP eşikleri, rozet kriterleri) düzenleyebilir
- Mentor ilişkisi tanımlayabilir
- Öğrenciye sertifika oluşturabilir
- Grafik ve raporları görüntüleyebilir, CSV export edebilir
- Platform ayarlarını güncelleyebilir, DB yedeği alabilir

**Öğrenci:**
- Kayıtlı kurslarını ve ilerleme yüzdesini görüntüleyebilir
- XP, level ve rozetlerini takip edebilir
- Mentor bilgisini görüntüleyebilir
- Sertifikalarını görüntüleyebilir

---

## 🖥️ Teknolojiler

| Teknoloji | Kullanım Alanı |
|-----------|----------------|
| Python 3.9+ | Ana programlama dili |
| PyQt5 | GUI Framework |
| SQLite3 | Veritabanı yönetimi (`@contextmanager`) |
| QPainter | Tüm grafikler (Bar, Pie, Line) |
| hashlib (SHA256) | Şifre güvenliği |
| csv | Rapor export |
