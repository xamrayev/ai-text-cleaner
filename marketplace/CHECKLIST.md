# Marketplace'ga yuborish: tartib va ro'yxat

## 1. Joylashdan oldin

- [ ] Doimiy HTTPS domenni tanlang (masalan, `https://aitextclean.ubs.uz` yoki GitHub Pages/Netlify/Vercel). Domen keyin o'zgarmasligi kerak, chunki manifest unga bog'langan.
- [ ] Ommaviy yordam pochtasini tanlang (masalan, `support@...`). U uchala sahifada va tekshiruvchilar uchun izohda ko'rinadi.
- [ ] Yig'ing:
  ```bash
  npm test
  npm run build -- --url https://SIZNING-DOMEN --email support@SIZNING-DOMEN
  ```
- [ ] `dist/site/` ichidagi hamma narsani domenning ildiziga yuklang. `https://SIZNING-DOMEN/taskpane.html` ochilishi kerak.
- [ ] Brauzerda tekshiring: `/privacy.html`, `/terms.html`, `/support.html`, `/assets/icon-64.png`.
- [ ] `npm run validate:prod`: «The manifest is valid» chiqishi kerak. Tekshiruv icon va support URL'larini haqiqatan ochib ko'radi.

## 2. Uch platformada sinov

`dist/manifest.xml` ni sideload qiling va [certification-notes.md](certification-notes.md) dagi qadamlarni bajaring:

- [ ] **Word Online:** *Вставка → Надстройки → Мои надстройки → Загрузить мою надстройку*.
- [ ] **Windows:** umumiy papka orqali (README'ga qarang) yoki `npx office-addin-debugging start dist/manifest.xml desktop`.
- [ ] **Mac:** manifestni `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/` papkasiga nusxalang va Word'ni qayta oching.
- [ ] Har bir platformada: lentadagi 3 tugma bildirishnoma chiqaradimi, panel ochiladimi, EN/RU/UZ almashadimi?
- [ ] Office tilini ruschaga o'zgartirib ko'ring: lenta va panel ruscha bo'lishi kerak.

## 3. Rasmlar (hozircha yo'q, tayyorlash kerak)

- [ ] **Do'kon logotipi:** kvadrat PNG. Partner Center hozirgi o'lchamni ko'rsatadi, odatda 216×216 dan 350×350 gacha (300×300 qulay). `web/assets/icon-128.png` ni cho'zmang, asl vektordan qayta eksport qiling.
- [ ] **Skrinshotlar:** 1366×768 PNG, 1–5 ta. Tavsiya: (1) sariq AI izlari va panel ro'yxati, (2) tozalangandan keyin yashil joylar va natijalar jadvali, (3) Matn rejimi, (4) qoidalar ro'yxati, (5) lenta tugmalari va bildirishnoma. Inglizcha interfeysda oling.
- [ ] (Ixtiyoriy) qisqa video, YouTube havolasi.

## 4. Partner Center

- [ ] Ro'yxatdan o'tish: Partner Center → Microsoft 365 and Copilot dasturi. Shu yerda aniqlang: ro'yxatdan o'tish to'lovi bormi, jismoniy shaxs sifatida nashr qilish mumkinmi, O'zbekiston ro'yxatda bormi. Bo'lmasa, UBS nomidan nashr qilish mumkin, bunda manifestdagi `ProviderName` va sahifalardagi nashriyotchi nomini o'zgartirish kerak bo'ladi.
- [ ] Yangi offer → *Office add-in*.
- [ ] **Product setup:** `dist/manifest.xml` ni yuklang.
- [ ] **Properties:** kategoriyalar, maxfiylik siyosati URL (`/privacy.html`), yordam URL (`/support.html`). Litsenziya uchun Microsoft standart shartnomasini tanlang yoki o'zingizniki uchun `/terms.html` ni ko'rsating.
- [ ] **Marketplace listings:** [listing.md](listing.md) dagi en-US va ru-RU matnlarini kiriting, logotip va skrinshotlarni yuklang. Bu yerda tanlangan tillar manifestda bo'lishi shart: en-US standart til, ru-RU esa Override sifatida bor.
- [ ] **Availability:** bepul.
- [ ] **Notes for certification:** [certification-notes.md](certification-notes.md) matnini qo'ying, `{{SUPPORT_EMAIL}}` ni almashtiring.
- [ ] Yuboring. Rad etilsa, hisobotdagi har bir bandni tuzatib, qayta yuboring.

## 5. Keyin

- Har bir yangilanishda manifestdagi `<Version>` ni oshiring (hozir 1.1.0.0) va qayta yuboring. Faqat `web/` fayllari o'zgarsa, ularni hostingda yangilashning o'zi yetadi va qayta tekshiruv kerak emas.
- Maxfiylik siyosati yoki shartlar o'zgarsa, sahifalardagi «Effective date» sanasini yangilang.

## Ma'lum cheklovlar

- O'zbek tili (`uz-Latn-UZ`) manifestda bor, shuning uchun interfeys o'zbekcha Office'da o'zbekcha ko'rinadi. Lekin Marketplace sahifasi tillari ro'yxatida o'zbekcha bo'lmasligi mumkin. Bunda listing faqat en/ru bo'ladi.
- Bildirishnomalar Office Dialog oynasida chiqadi, chunki Word'da function command uchun boshqa bildirishnoma API yo'q. Ular 6 soniyadan keyin o'zi yopiladi.
