
from pptx import Presentation
from pptx.util import Inches
import openpyxl
import os

# Путь к Excel и PowerPoint
excel_path = r"D:/VSCode/AutoPowerPoint.xlsx"
pptx_path = r"D:/VSCode/Emoji.pptx"

# Загружаем Excel
wb = openpyxl.load_workbook(excel_path)
ws = wb.active

# Загружаем презентацию
prs = Presentation(pptx_path)

# Пробегаемся по строкам Excel
for row in ws.iter_rows(min_row=2, values_only=True):  # пропускаем заголовки
    image_path, slide_num  = row
    if not os.path.isfile(image_path):
        print(f"Файл не найден: {image_path}")
        continue
    if slide_num > len(prs.slides):
        print(f"Нет слайда {slide_num} в презентации")
        continue
    
    slide = prs.slides[slide_num - 1]
    
    # Вставка картинки (можно настроить позицию и размер)
    slide.shapes.add_picture(image_path, Inches(1), Inches(1), width=Inches(4.5))

# Сохраняем презентацию
prs.save("presentation_with_images.pptx")
print("✅ Готово!")
