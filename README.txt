HISOB KO'CHIRMASI TAHLILI — qaytib chiqmagan tushumlarni aniqlash
==================================================================

1. O'RNATISH (bir marta)
   - Kompyuterda Python 3.9+ bo'lishi kerak: https://www.python.org/downloads/
     O'rnatishda "Add python.exe to PATH" belgisini albatta qo'ying.
   - Arxivni istalgan papkaga oching.

2. ISHGA TUSHIRISH
   - "ishga_tushirish.bat" faylini ikki marta bosing.
     Birinchi safar kerakli kutubxonalar internetdan avtomatik o'rnatiladi.
   - Oynada "Tanlash..." orqali bank tizimidan olingan
     "Справка о работе счета" faylini (.xlsb / .xlsx) tanlang.
   - "Tahlil qilish" tugmasini bosing. Natija Excel fayli saqlanadi va ochiladi.

   Buyruq satridan:  ishga_tushirish.bat "D:\fayl.xlsb" "D:\natija.xlsx"

3. NATIJA FAYLI VARAQLARI
   - Xulosa                 — umumiy ko'rsatkichlar, tekshiruv, toifalar, oylar
   - Qaytmagan tushumlar    — hisobga tushib, chiqim bilan bog'lanmagan tushumlar
   - To'lovchilar kesimida  — qaytmagan tushumlar MIB bo'limlari bo'yicha
   - Bog'lanmagan chiqimlar — hech bir tushumga bog'lanmagan chiqimlar
   - Moslangan juftliklar   — qaysi tushum qaysi chiqim bilan, qaysi usulda yopilgan

4. MOSLASHTIRISH MANTIG'I (tartib bo'yicha)
   1) Summa teng + to'lov maqsadi matni chiqim matnida takrorlangan
   2) Summa teng + umumiy ijro hujjati / PINFL / sud ishi raqami
   3) Summa teng + chiqim oluvchisining familiyasi tushum matnida bor
   4) Summa teng, chiqim tushumdan keyin 31 kun ichida
   5) Bir nechta tushum yig'indisi = bitta chiqim (bir shaxs)
   6) Summa deyarli teng (≤0,5%) + shaxs bir xil
   7) Storno — chiqim keyinroq hisobga qaytib tushgan
   8) Bitta tushum bir kunda ko'p hisobga taqsimlangan
   9) Chiqim = shu shaxs avvalgi tushumlarining FIFO yig'indisi
   Chiqim har doim tushumdan keyin (yoki o'sha kuni) bo'lishi shart.

   Tekshiruv: (qaytmagan tushumlar − bog'lanmagan chiqimlar) ≈ (chiqish qoldig'i − kirish qoldig'i)

DIQQAT: natija avtomatik moslashtirish natijasidir. "Bog'liq bo'lishi mumkin
bo'lgan chiqim qatori" ustuni to'ldirilgan qatorlarni qo'lda tekshirib chiqing.
