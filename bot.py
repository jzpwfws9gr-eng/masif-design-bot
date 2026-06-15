import os
from PIL import Image, ImageDraw, ImageFont
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes

TOKEN = os.getenv("BOT_TOKEN")

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "ارسل كذا:\n/group A\nاسبانيا 3\nالسعودية 1\nالأوروغواي 1\nنيوزيلندا 0"
    )

async def group(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.splitlines()
    group_name = text[0].replace("/group", "").strip() or "A"
    teams = []

    for line in text[1:5]:
        parts = line.rsplit(" ", 1)
        if len(parts) == 2:
            teams.append((parts[0], parts[1]))

    img = Image.new("RGB", (1080, 1350), (22, 10, 45))
    draw = ImageDraw.Draw(img)

    try:
        font_big = ImageFont.truetype("DejaVuSans-Bold.ttf", 72)
        font_med = ImageFont.truetype("DejaVuSans-Bold.ttf", 48)
        font_small = ImageFont.truetype("DejaVuSans-Bold.ttf", 42)
    except:
        font_big = font_med = font_small = ImageFont.load_default()

    draw.text((540, 120), "ترتيب المجموعة", fill="white", font=font_big, anchor="mm")
    draw.text((540, 220), group_name, fill=(255, 210, 80), font=font_big, anchor="mm")

    y = 380
    for i, (name, pts) in enumerate(teams, start=1):
        draw.rounded_rectangle((80, y-45, 1000, y+55), radius=25, fill=(45, 20, 85))
        draw.text((950, y), str(i), fill=(255, 210, 80), font=font_med, anchor="rm")
        draw.text((760, y), name, fill="white", font=font_small, anchor="rm")
        draw.text((130, y), pts, fill="white", font=font_med, anchor="lm")
        y += 140

    draw.text((540, 1200), "MONDIAL ALMASIF 2026", fill=(255, 210, 80), font=font_small, anchor="mm")

    path = "group.png"
    img.save(path)

    await update.message.reply_photo(photo=open(path, "rb"))

def main():
    app = ApplicationBuilder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("group", group))
    app.run_polling()

if __name__ == "__main__":
    main()