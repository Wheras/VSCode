from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
import openpyxl
import os

# Пути
excel_path = r"D:/VSCode/AutoPowerPoint.xlsx"
pptx_path = r"D:/VSCode/Emoji.pptx"
background_path = r"D:\VSCode\background.jpg"
# Загружаем Excel
wb = openpyxl.load_workbook(excel_path)
ws = wb.active

# Загружаем или создаём презентацию
prs = Presentation()
prs.slide_width = Inches(16)
prs.slide_height = Inches(9)

# Шаблон слайда (можно менять)
blank_layout = prs.slide_layouts[6]  # пустой слайд

# Координаты и размер картинки
img_left = Inches(2)
img_top = Inches(1.5)
img_width = Inches(6)  # можно менять
img_height = Inches(4)  # если нужно фиксированное соотношение

# Пробегаемся по строкам Excel
for row in ws.iter_rows(min_row=2, values_only=True):  # пропускаем заголовки
    image_path, slide_num  = row

    if not os.path.isfile(image_path):
        print(f"❌ Картинка не найдена: {image_path}")
        continue

    # Создаём слайд
    slide = prs.slides.add_slide(blank_layout)

    # Установка фонового изображения
    if background_path and os.path.isfile(background_path):
        slide.shapes.add_picture(background_path, 0, 0, width=prs.slide_width, height=prs.slide_height)
    else:
        print(f"⚠ Фон не найден или пуст: {background_path}")

    # Вставка картинки
    pic = slide.shapes.add_picture(image_path, img_left, img_top, width=img_width, height=img_height)

# Сохраняем презентацию
prs.save("presentation_with_images.pptx")
print("✅ Готово!")
