from flask import Flask, render_template, request, send_file, redirect, url_for, flash
from bs4 import BeautifulSoup
from openpyxl import Workbook
import requests
import json
from io import BytesIO
from urllib.parse import urljoin

app = Flask(__name__)
app.secret_key = "my_super_secret_key_12345"  # можешь заменить на свой сгенерированный ключ

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

    # --- Пробуем найти JSON-LD ---
    script_tag = soup.find("script", type="application/ld+json")
    if script_tag:
        try:
            data = json.loads(script_tag.string)
            if "itemListElement" in data:
                for item in data["itemListElement"]:
                    name = item.get("name", "Без названия").strip()
                    description = item.get("description", "").replace("\n", " ").strip()
                    price = int(item.get("price", "0").strip()) if item.get("price") else 0
                    new_price = int(price * 1.2)

                    image_url = None
                    if "image" in item:
                        if isinstance(item["image"], str):
                            image_url = urljoin(url, item["image"])
                        elif isinstance(item["image"], list) and len(item["image"]) > 0:
                            image_url = urljoin(url, item["image"][0])

                    results.append({
                        "name": name,
                        "description": description,
                        "original_price": price,
                        "new_price": new_price,
                        "image_url": image_url
                    })
        except Exception as e:
            print(f"⚠ Ошибка JSON-LD: {e}")

    # --- Если JSON-LD нет, парсим HTML ---
    if not results:
        menupoints = soup.find_all("div", class_="menupoint")
        for block in menupoints:
            name_tag = block.find("div", class_="menuline")
            name = name_tag.get_text(strip=True) if name_tag else "Без названия"

            desc_tag = block.find("div", class_="description")
            description = desc_tag.get_text(" ", strip=True) if desc_tag else ""

            price_block = block.find("div", class_="price")
            price_text = price_block.get_text(" ", strip=True) if price_block else "0"
            price_digits = "".join([c for c in price_text if c.isdigit()])
            price = int(price_digits) if price_digits else 0
            new_price = int(price * 1.2)

            # === Ищем фото ===
            img_tag = block.find("img")
            image_url = None
            if img_tag and img_tag.get("src"):
                image_url = urljoin(url, img_tag["src"])

            results.append({
                "name": name,
                "description": description,
                "original_price": price,
                "new_price": new_price,
                "image_url": image_url
            })

    return results


# ==============================
# Сохранение в Excel и TXT
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
            f"{item['original_price']}₽",
            f"{item['new_price']}₽",
            item.get("image_url") or ""
        ])
    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream


def save_to_txt(data):
    text = ""
    for item in data:
        text += f"{item['name']} — {item['original_price']}₽ → {item['new_price']}₽\n"
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
