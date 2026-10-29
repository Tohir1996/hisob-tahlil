# Hisob Ko'chirmasi Tahlili (Bank Statement Analyzer)

Bank hisob raqami ko'chirmasi (*«Справка о работе счета»*) fayllarini (`.xlsb` / `.xlsx` / `.xls`) tahlil qilib, **hisobga kelib tushgan, ammo chiqim bilan yopilmagan (qaytib chiqmagan)** tushumlarni avtomatik aniqlash dasturi.

---

## 🚀 Boshqa kompyuterda o'rnatish va ishga tushirish

### 1. Repozitoriyni yuklab olish
Terminal (PowerShell / CMD) orqali:
```bash
git clone https://github.com/Tohir1996/hisob-tahlil.git
cd hisob-tahlil
```
*(Yoki GitHub saytidan **Code -> Download ZIP** tugmasi orqali yuklab olib, arxivdan chiqaring).*

### 2. Ishga tushirish (Windows)
Faqat **`ishga_tushirish.bat`** faylini ikki marta bosing:
- Kompyuterda Python bor-yo'qligi tekshiriladi;
- Birinchi marta ishga tushganda barcha kerakli kutubxonalar (`pandas`, `openpyxl`, `pyxlsb`, `xlrd`) avtomatik o'rnatiladi;
- Qulay grafik oyna (GUI) ochiladi.

> **Eslatma:** Kompyuterda [Python 3.9+](https://www.python.org/downloads/) o'rnatilgan bo'lishi va o'rnatish jarayonida **"Add python.exe to PATH"** belgilangan bo'lishi lozim.

---

## 🖥 Foydalanish

### Grafik interfeys (GUI) orqali:
1. **Ko'chirma fayli** qatorida `Tanlash...` tugmasini bosib bank faylini (`.xlsb` yoki `.xlsx`) tanlang;
2. **`▶ Tahlil qilish`** tugmasini bosing;
3. Tahlil tugagach, hosil bo'lgan Excel faylni to'g'ridan-to'g'ri ochish taklif qilinadi.

### Buyruqlar satri (CLI) orqali:
```bash
python hisob_tahlil.py "D:\b2r74D8.xlsb" "D:\natija.xlsx"
```

---

## 📊 Natijaviy Excel hisobot tarkibi

Dastur yakunida batafsil tuzilgan Excel fayl yaratadi:
1. **Xulosa** — Hisob raqami, aylanma, balans tekshiruvi, qaytmagan tushumlar toifalari va oylar kesimi;
2. **Qaytmagan tushumlar** — Barcha chiqim bilan bog'lanmagan tushumlar (sana, to'lovchi, summa, PINFL, necha kun turgani, ehtimoliy qisman chiqimlar);
3. **To'lovchilar kesimida** — MIB bo'limlari va boshqa to'lovchilar bo'yicha yig'ma jadval;
4. **Bog'lanmagan chiqimlar** — Hech qanday tushumga bog'lanmagan chiqim operatsiyalari;
5. **Moslangan juftliklar** — Qaysi tushum qaysi chiqim bilan va qaysi mantiq asosida bog'langanligi.

---

## 🧠 Tahlil va moslashtirish mantig'i

Tushum va chiqimlar 9 bosqichli intellektual algoritm asosida solishtiriladi:
1. **1-bosqich:** Summa teng + to'lov maqsadi matni chiqim matnida to'liq takrorlangan;
2. **2-bosqich:** Summa teng + umumiy ijro hujjati / PINFL / sud ishi raqami mavjud;
3. **3-bosqich:** Summa teng + chiqim oluvchisining familiyasi tushum matnida uchraydi;
4. **4-bosqich:** Summa teng + chiqim tushumdan keyin 31 kun ichida amalga oshirilgan;
5. **5-bosqich:** Bir nechta tushumlar yig'indisi bitta chiqimga teng (bir xil shaxs bo'yicha);
6. **6-bosqich:** Summa deyarli teng (farq ≤0.5%) + qarzdor shaxsi bir xil;
7. **7-bosqich:** Storno operatsiyalari (chiqim keyinroq hisobga qaytarilgan holatlar);
8. **8-bosqich:** Bitta yirik tushum bir kunda ko'plab hisoblarga taqsimlangan holatlar;
9. **9-bosqich:** Bir shaxs bo'yicha avvalgi tushumlarning FIFO yig'indisi.
