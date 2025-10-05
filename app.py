from flask import Flask, render_template, request, send_file, redirect, url_for, flash
from bs4 import BeautifulSoup
from openpyxl import Workbook
import requests
import json
from io import BytesIO
from urllib.parse import urljoin
import re

app = Flask(__name__)
app.secret_key = "my_super_secret_key_12345"  # можешь заменить на свой ключ

# ==============================
# Функция парсинга
# ==============================
def parse_rolls_page(url: str):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                      "AppleWebKit/537.36 (KHTML, like Gecko) "
                      "Chrome/128.0.0.0 Safari/537.36"
    }
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, "html.parser")

    results = []

    menupoints = soup.find_all("div", class_="menupoint")

    for block in menupoints:
        # ===== Название =====
        name_tag = block.find("div", class_="menuline")
        name = name_tag.find("a").get_text(strip=True) if name_tag and name_tag.find("a") else "Без названия"

        # ===== Описание =====
        desc_tag = block.find("div", class_="description")
        description = desc_tag.get_text(" ", strip=True) if desc_tag else ""

        # ===== Цена =====
        price_block = block.find("div", class_="price")
        price_text = ""
        if price_block:
            # Удаляем все span (перечёркнутые цены)
            for span in price_block.find_all("span"):
                span.extract()

            # Берём оставшийся текст, который содержит актуальную цену
            price_text = price_block.get_text(" ", strip=True)

        # Извлечение рублей и копеек
        price = 0.0
        price_match = re.search(r"(\d+)\s*р\.?\s*(\d*)\s*к\.?", price_text)
        if price_match:
            rubles = int(price_match.group(1))
            kopeks = int(price_match.group(2)) if price_match.group(2) else 0
            price = rubles + kopeks / 100
        else:
            price_digits = "".join([c for c in price_text if c.isdigit()])
            price = int(price_digits) if price_digits else 0

        new_price = round(price * 1.2, 2)

        # ===== Фото =====
        img_tag = block.find("img")
        image_url = urljoin(url, img_tag["src"]) if img_tag and img_tag.get("src") else None

        results.append({
            "name": name,
            "description": description,
            "original_price": price,
            "new_price": new_price,
            "image_url": image_url
        })

    return results


# ==============================
# Сохранение в Excel
# ==============================
def save_to_excel(data):
    wb = Workbook()
    ws = wb.active
    ws.title = "Сеты роллов"
    ws.append(["Название", "Описание", "Цена (оригинал)", "Цена (+20%)", "Ссылка на фото"])
    for item in data:
        ws.append([
            item["name"],
            item["description"],
            f"{item['original_price']:.2f}₽",
            f"{item['new_price']:.2f}₽",
            item.get("image_url") or ""
        ])
    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream


# ==============================
# Сохранение в TXT
# ==============================
def save_to_txt(data):
    text = ""
    for item in data:
        text += f"{item['name']} — {item['original_price']:.2f}₽ → {item['new_price']:.2f}₽\n"
        text += f"Описание: {item['description']}\n"
        if item.get("image_url"):
            text += f"Фото: {item['image_url']}\n"
        text += "\n"
    return BytesIO(text.encode("utf-8"))


# ==============================
# Flask маршруты
# ==============================
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        url = request.form.get("url")
        if not url:
            flash("Введите ссылку!", "error")
            return redirect(url_for("index"))
        try:
            rolls = parse_rolls_page(url)
            if not rolls:
                flash("Не удалось найти данные на странице.", "error")
                return redirect(url_for("index"))
            return render_template("results.html", rolls=rolls, url=url)
        except Exception as e:
            flash(f"Ошибка парсинга: {e}", "error")
            return redirect(url_for("index"))
    return render_template("index.html")


@app.route("/download/excel", methods=["POST"])
def download_excel():
    data = json.loads(request.form["data"])
    stream = save_to_excel(data)
    return send_file(stream, as_attachment=True, download_name="rolls.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/download/txt", methods=["POST"])
def download_txt():
    data = json.loads(request.form["data"])
    stream = save_to_txt(data)
    return send_file(stream, as_attachment=True, download_name="rolls.txt",
                     mimetype="text/plain")


# ==============================
# Запуск
# ==============================
if __name__ == "__main__":
    app.run(debug=True)
