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

Admin paneldagi `📊 Excel` tugmasi barcha arizalarni `.xlsx` fayl ko'rinishida yuboradi.
Excel eksportida nomzod rasmi ham jadval ichida ko'rinadi.

Admin panelda arizalarni status bo'yicha ko'rish, ism/telefon/filial bo'yicha qidirish,
filial bo'yicha filter qilish va rad etishda sabab yozish imkoniyati bor.

Foydalanuvchi tug'ilgan sanani kiritadi, bot yoshni avtomatik hisoblaydi. Rasm yuborishdan
oldin ariza ma'lumotlarini tekshirib, kerak bo'lsa tahrirlashi mumkin.
