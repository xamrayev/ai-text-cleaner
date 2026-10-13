# AI Text Cleaner (AI Matn Tozalagich) — Word nadstroykasi

ChatGPT, Claude, Gemini chiqargan matndagi ko‘rinmas belgilar, jingalak qo‘shtirnoqlar, tirelar va markdown qoldiqlarini **to‘g‘ridan-to‘g‘ri Word hujjatida** tozalaydi. aitextclean.com g‘oyasi asosida, lekin Word uchun moslashtirilgan.

## Imkoniyatlar

- **AI izlarini ko‘rsatish** — iz bor so‘zlar hujjatda **sariq** bo‘yaladi, panelda esa ro‘yxat chiqadi: qaysi so‘zda, qanday iz (ko‘rinmas belgilar ¤ ⍽ ␣ bilan ko‘rsatiladi). Ro‘yxatdagi joy bosilsa, Word shu yerga o‘tadi.
- **Tozalash** — tozalangan joylar **yashil** rangga o‘tadi (avval sariq bo‘lganlari ham).
- **Indikator va «Ranglarni o‘chirish»** — hujjatda nechta rangli joy borligini ko‘rsatadi. Tugma bosilsa, faqat nadstroyka qo‘ygan ranglar olinadi, sizning avvalgi bo‘yoqlaringiz tiklanadi (joylar yashirin xatcho‘plar bilan eslab qolinadi).
- **Hujjat rejimi** — belgilangan qism yoki butun hujjatni tahlil qilish va tozalash. Matn qayta yozilmaydi: faqat kerakli belgilar qidirib almashtiriladi, shuning uchun shrift, rang, havolalar, jadval va rasmlar saqlanadi.
- **Matn rejimi** — matnni panelga qo‘yib tozalash, nusxalash yoki kursor turgan joyga qo‘yish.
- **Lenta tugmalari:** «Tozalagich» (panel), «AI izlari», «Tez tozalash», «Ranglarni o‘chirish» — oxirgi uchtasi panelni ochmasdan ishlaydi.
- **Emoji buzilmaydi:** 👨‍👩‍👧, ❤️, 1️⃣ va bayroqlar ichidagi ZWJ, variation selector va teglar saqlanadi. Matnga yashirib qo‘yilgan Unicode teglar (U+E0000–E007F), boshqa joylardagi variation selector va bo‘sh ko‘rinadigan to‘ldiruvchi belgilar (U+3164, U+2800 va boshqalar) o‘chiriladi.
- **10 ta qoida**, har biri alohida yoqiladi: yashirin belgilar, NBSP, tirelar, qo‘shtirnoqlar, `…`, qator oxiridagi bo‘shliqlar, `*`, `#` sarlavhalar, o‘xshash harflar, NFKC.
- **O‘zbek tiliga moslashuv:** `o‘`, `g‘` va tutuq belgisidagi (`ma’no`) apostroflar saqlanadi, `«»` qo‘shtirnoqlarga standart holatda tegilmaydi (xohlasangiz, ««» qo‘shtirnoqlarni ham almashtirish» sozlamasini yoqing). O‘xshash harflar qoidasi faqat *aralash* so‘zlarni (lotin + kirill) tuzatadi, sof kirill matnni buzmaydi.
- **O‘zgarishlarni kuzatish rejimi** — har bir tuzatish Review bo‘limida ko‘rinadi (WordApi 1.4).
- **Lenta tugmalari natijani ko‘rsatadi** — kichik oynada nechta o‘zgarish yoki joy topilgani chiqadi (6 soniyada o‘zi yopiladi).
- **Uch tilli interfeys:** inglizcha, ruscha, o‘zbekcha. Office tili bo‘yicha avtomatik tanlanadi, panel yuqorisidagi EN/RU/UZ orqali almashtiriladi.
- Hech qanday server yoki API yo‘q, matn kompyuterdan chiqmaydi.

## Tuzilma

```
manifest.xml         nadstroyka manifesti (localhost, en + ru-RU/uz-Latn-UZ tarjimalari)
web/cleaner.js       tozalash yadrosi (Word'ga bog'liq emas)
web/i18n.js          interfeys tarjimalari (en/ru/uz) va tilni aniqlash
web/word-engine.js   Word hujjatida qidirib-almashtirish
web/taskpane.*       panel interfeysi
web/commands.*       lenta tugmalari (Tez tozalash, AI izlari, Ranglarni o'chirish)
web/notice.html      lenta tugmalari natijasi chiqadigan oyna
web/privacy.html, terms.html, support.html   maxfiylik, shartlar, yordam (3 tilda)
scripts/build.js     Marketplace uchun yig'ish (dist/)
marketplace/         listing matni, tekshiruvchilar uchun izoh, yuborish ro'yxati
web/settings.js      sozlamalarni saqlash
test/run-tests.js    testlar (Word API imitatsiyasi bilan)
```

## Ishga tushirish (dasturchi rejimi, Windows/Mac)

Node.js 18+ kerak.

```bash
npm install
npm run certs      # localhost uchun ishonchli HTTPS sertifikat (bir marta)
npm start          # serverni ishga tushiradi va Word'ga nadstroykani yuklaydi
```

Word ochiladi, **Bosh sahifa (Главная)** lentasida «AI Tozalagich» guruhi paydo bo‘ladi. To‘xtatish: `npm run stop`. Testlar: `npm test`.

**Word Online'da sinash:** `npm run serve` ni ishga tushiring, so‘ng Word Online'da *Вставка → Надстройки → Мои надстройки → Загрузить мою надстройку* orqali `manifest.xml` ni yuklang.

## GitHub Pages

Sayt: https://xamrayev.github.io/ai-text-cleaner/ — `main` ga har bir push'da [.github/workflows/pages.yml](.github/workflows/pages.yml) testlarni ishga tushiradi, `scripts/build.js` bilan yig'adi va joylaydi. Aloqa pochtasi repo o'zgaruvchisidan olinadi (`SUPPORT_EMAIL`, *Settings → Secrets and variables → Actions → Variables*).

Tayyor manifest: https://xamrayev.github.io/ai-text-cleaner/manifest.xml — uni yuklab olib, Word'ga «Загрузить мою надстройку» orqali qo'shish mumkin.

## Boshqalarga tarqatish

1. Yig'ing: `npm run build -- --url https://SIZNING-DOMEN --email support@SIZNING-DOMEN`. Skript `dist/site/` (hosting uchun) va `dist/manifest.xml` ni yaratadi, barcha `localhost` va `{{SUPPORT_EMAIL}}` almashtiriladi.
2. `dist/site/` ni domen ildiziga joylang (GitHub Pages, Netlify, Vercel, universitet serveri). So'ng `npm run validate:prod` ni ishga tushiring.
3. `dist/manifest.xml` ni tarqating:
   - **Tashkilot ichida (UBS):** Microsoft 365 admin markazi → *Integrated apps* → manifestni yuklash. Hamma xodimlarning Word'ida avtomatik paydo bo‘ladi.
   - **Bitta kompyuterda (Windows):** manifestni umumiy tarmoq papkasiga qo‘ying, Word → *Параметры → Центр управления безопасностью → Каталоги доверенных надстроек* ga shu papkani qo‘shing.
   - **Hammaga ochiq:** Microsoft Marketplace (Partner Center). To'liq tartib: [marketplace/CHECKLIST.md](marketplace/CHECKLIST.md).

## Cheklovlar

- Maydonlar (field), matn qutilari va izohlar ichidagi belgilarni Word qidiruvi har doim ham topmaydi. Bunday holat bo‘lsa, panel nechta belgi qolganini aytadi.
- Belgilangan qism bilan ishlaganda, sarlavha va oxirgi bo‘shliq qoidalari butun paragrafga qo‘llanadi.
- Kuzatish rejimida hisoblagich aniq bo‘lmasligi mumkin.
- Ranglash so‘z darajasida ishlaydi (ko‘rinmas belgining o‘zini bo‘yab bo‘lmaydi, shuning uchun u turgan so‘z bo‘yaladi). WordApi 1.3 kerak; ranglarni aniq tiklash uchun WordApi 1.4 (Microsoft 365, Word Online, Word 2021+). Eski versiyalarda «Ranglarni o‘chirish» faqat sof sariq/yashil bo‘yoqlarni oladi.
