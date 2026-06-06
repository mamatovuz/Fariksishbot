# fariksaloqabot

Telegram orqali ariza qabul qiladigan va admin tasdiqlagandan keyin kanalga joylaydigan bot.

Ma'lumotlar lokal `bot.db` SQLite bazasida saqlanadi va `applications.xlsx`
Excel fayliga yozib boriladi.

Localda bot polling rejimida ishlaydi. Railway kabi hostlarda `WEBHOOK_URL` yoki
`RAILWAY_PUBLIC_DOMAIN` bo'lsa webhook rejimi avtomatik yoqiladi.

## Ishga tushirish

```powershell
pip install -r requirements.txt
python main.py
```

`Conflict: terminated by other getUpdates request` chiqsa, bir xil token bilan bot ikki
joyda ishlayapti. Kompyuterdagi eski `python main.py` jarayonini to'xtating yoki hostdagi
replica sonini 1 taga tushiring.

Admin paneldagi `📊 Excel` tugmasi faqat tasdiqlangan arizalarni `.xlsx` fayl
ko'rinishida yuboradi. Excel eksportida nomzod rasmi ham jadval ichida ko'rinadi,
Telegram `file_id` qiymati Excelga yozilmaydi.

Admin panelda arizalarni status bo'yicha ko'rish, ism/telefon/filial bo'yicha qidirish,
filial bo'yicha filter qilish va rad etishda sabab yozish imkoniyati bor.
Bir arizani faqat bitta admin yakunlay oladi: kim birinchi tasdiqlasa yoki rad etsa,
qolgan adminlardagi inline tugmalar ishlamaydi va qaror haqida xabar boradi.
Admin paneldagi `📣 Xabar yuborish` tugmasi orqali barcha foydalanuvchilarga matn,
rasm, video yoki GIF yuborish mumkin. Media captionidagi matn ham birga boradi.

Foydalanuvchi tug'ilgan sanani kiritadi, bot yoshni avtomatik hisoblaydi. Rasm yuborishdan
oldin ariza ma'lumotlarini tekshirib, kerak bo'lsa tahrirlashi mumkin.

Ariza boshlanishidan oldin foydalanuvchi `@fariks01` kanaliga obuna bo'lishi kerak.
Obunani tekshirish ishlashi uchun bot kanalga admin qilib qo'yilishi kerak.

Ma'lumot darajasidan keyin foydalanuvchi yo'nalishni tanlaydi: Admin, O'qituvchi yoki
O'qituvchi yordamchi. O'qituvchi yo'nalishlarida mutaxassislik va sertifikat rasmi
so'raladi.
