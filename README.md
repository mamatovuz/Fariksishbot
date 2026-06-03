# fariksaloqabot

Telegram orqali ariza qabul qiladigan va admin tasdiqlagandan keyin kanalga joylaydigan bot.

Ma'lumotlar lokal `bot.db` SQLite bazasida saqlanadi.

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
