import os
from PIL import Image, ImageDraw, ImageFont
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes

TOKEN = os.getenv("BOT_TOKEN")


def get_font(size):
    try:
        return ImageFont.truetype(
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            size
        )
    except:
        return ImageFont.load_default()


async def group(update: Update, context: ContextTypes.DEFAULT_TYPE):
    lines = update.message.text.splitlines()
    group_name = lines[0].replace("/group", "").strip() or "A"

    teams = []
    for line in lines[1:5]:
        parts = line.rsplit(" ", 1)
        if len(parts) == 2:
            name = parts[0].strip()
            pts = parts[1].strip()
            teams.append((name, pts))

    img = Image.new("RGB", (1080, 1080), (18, 8, 45))
    draw = ImageDraw.Draw(img)

    title_font = get_font(60)
    group_font = get_font(80)
    row_font = get_font(52)
    pts_font = get_font(55)
    small_font = get_font(28)

    gold = (218, 172, 62)
    purple = (55, 24, 100)
    dark = (18, 8, 45)

    draw.text(
        (540, 70),
        "GROUP STANDINGS",
        fill="white",
        font=title_font,
        anchor="mm"
    )

    draw.rounded_rectangle(
        (390, 110, 690, 200),
        radius=25,
        fill=gold
    )

    draw.text(
        (540, 155),
        group_name.upper(),
        fill=dark,
        font=group_font,
        anchor="mm"
    )

    y = 290

    for i in range(4):
        name, pts = teams[i] if i < len(teams) else ("-", "0")

        draw.rounded_rectangle(
            (80, y - 40, 1000, y + 50),
            radius=22,
            fill=purple
        )

        draw.rounded_rectangle(
            (850, y - 40, 1000, y + 50),
            radius=22,
            fill=gold
        )

        draw.text(
            (925, y),
            str(i + 1),
            fill=dark,
            font=pts_font,
            anchor="mm"
        )

        draw.text(
            (760, y),
            name,
            fill="white",
            font=row_font,
            anchor="rm"
        )

        draw.rounded_rectangle(
            (80, y - 40, 180, y + 50),
            radius=20,
            outline=gold,
            width=3
        )

        draw.text(
            (130, y),
            pts,
            fill="white",
            font=pts_font,
            anchor="mm"
        )

        y += 130

    draw.text(
        (540, 930),
        "POINTS",
        fill=gold,
        font=small_font,
        anchor="mm"
    )

    draw.text(
        (540, 980),
        "MONDIAL ALMASIF 2026",
        fill=gold,
        font=small_font,
        anchor="mm"
    )

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