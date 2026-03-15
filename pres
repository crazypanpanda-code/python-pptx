from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import Color
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# Создаем презентацию
prs = Presentation()

# Цвета
DARK_BG = Color.from_string('#0A0F1F')
NEON_BLUE = Color.from_string('#00D4FF')
NEON_RED = Color.from_string('#FF4757')
NEON_GREEN = Color.from_string('#2ED573')
NEON_ORANGE = Color.from_string('#FFA502')
NEON_YELLOW = Color.from_string('#FFFF00')
WHITE = Color.from_string('#FFFFFF')
GRAY = Color.from_string('#B0B8C4')

def set_slide_background(slide, color):
    """Устанавливает темный фон слайда"""
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color.rgb

def add_title_slide(prs, title_text, subtitle_text):
    """Создает титульный слайд"""
    slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(slide_layout)
    set_slide_background(slide, DARK_BG)
    
    # Заголовок
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(2), Inches(9), Inches(1.5))
    title_frame = title_box.text_frame
    title_para = title_frame.add_paragraph()
    title_para.text = title_text
    title_para.font.size = Pt(44)
    title_para.font.bold = True
    title_para.font.color.rgb = WHITE.rgb
    title_para.alignment = PP_ALIGN.CENTER
    
    # Подзаголовок
    subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(3.5), Inches(9), Inches(1))
    subtitle_frame = subtitle_box.text_frame
    subtitle_para = subtitle_frame.add_paragraph()
    subtitle_para.text = subtitle_text
    subtitle_para.font.size = Pt(24)
    subtitle_para.font.color.rgb = GRAY.rgb
    subtitle_para.alignment = PP_ALIGN.CENTER
    
    return slide

def add_section_slide(prs, title_text, icons_data):
    """Создает слайд с 4 секциями"""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    set_slide_background(slide, DARK_BG)
    
    # Заголовок
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(1))
    title_frame = title_box.text_frame
    title_para = title_frame.add_paragraph()
    title_para.text = title_text
    title_para.font.size = Pt(40)
    title_para.font.bold = True
    title_para.font.color.rgb = WHITE.rgb
    title_para.alignment = PP_ALIGN.CENTER
    
    # Создаем 4 квадрата с иконками
    colors = [NEON_RED, NEON_ORANGE, NEON_YELLOW, NEON_BLUE]
    positions = [(1, 1.5), (5.5, 1.5), (1, 4.5), (5.5, 4.5)]
    
    for i, (pos, color) in enumerate(zip(positions, colors)):
        # Квадрат
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(pos[0]), Inches(pos[1]),
            Inches(3.5), Inches(2.5)
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = color.rgb
        shape.line.color.rgb = color.rgb
        shape.line.width = Pt(3)
        
        # Текст иконки
        textbox = slide.shapes.add_textbox(
            Inches(pos[0] + 0.5), Inches(pos[1] + 0.8),
            Inches(2.5), Inches(1)
        )
        tf = textbox.text_frame
        p = tf.add_paragraph()
        p.text = icons_data[i]
        p.font.size = Pt(36)
        p.font.color.rgb = WHITE.rgb
        p.alignment = PP_ALIGN.CENTER
    
    return slide

def add_content_slide(prs, title_text, content_items, accent_color):
    """Создает слайд с контентом и маркерами"""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    set_slide_background(slide, DARK_BG)
    
    # Заголовок
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(1))
    title_frame = title_box.text_frame
    title_para = title_frame.add_paragraph()
    title_para.text = title_text
    title_para.font.size = Pt(36)
    title_para.font.bold = True
    title_para.font.color.rgb = accent_color.rgb
    title_para.alignment = PP_ALIGN.CENTER
    
    # Контент
    y_pos = 1.5
    for item in content_items:
        textbox = slide.shapes.add_textbox(Inches(0.5), Inches(y_pos), Inches(9), Inches(0.8))
        tf = textbox.text_frame
        p = tf.add_paragraph()
        p.text = f"• {item}"
        p.font.size = Pt(20)
        p.font.color.rgb = WHITE.rgb
        
        y_pos += 0.9
    
    return slide

# === СОЗДАНИЕ СЛАЙДОВ ===

# Слайд 1: Титульный
add_title_slide(
    prs,
    "КИБЕРУГРОЗЫ: КАРТА УЯЗВИМОСТЕЙ",
    "О чём должен думать админ каждую секунду"
)

# Слайд 2: 4 вектора атак
icons = ["👤️", "", "🌐", "❓"]
add_section_slide(prs, "4 ВЕКТОРА АТАК", icons)

# Слайд 3: Социальная инженерия
social_engineering_items = [
    "Фишинг (ссылки, письма)",
    "Претекстинг (легенда, звонок из 'безопасности банка')",
    "Кви про кво ('я из техподдержки, помогу настроить')",
    "Бейтинг ('бесплатный софт, скидки')",
    "",
    "Доверие стоит дороже антивируса"
]
add_content_slide(prs, "ВЕКТОР 1. СОЦИАЛЬНАЯ ИНЖЕНЕРИЯ", social_engineering_items, NEON_RED)

# Слайд 4: Вредоносное ПО
malware_items = [
    "Ransomware (Шифровальщики) — Цель: Данные. Исход: Бизнес встал.",
    "Spyware/Keyloggers — Цель: Пароли. Исход: Кража аккаунтов.",
    "Botnet — Цель: Ресурсы ПК. Исход: Сервер рассылает спам.",
    "Трояны — Цель: Задняя дверь. Исход: Полный контроль."
]
add_content_slide(prs, "ВЕКТОР 2. ВРЕДОНОСНОЕ ПО (MALWARE)", malware_items, NEON_ORANGE)

# Слайд 5: Сеть и сервисы
network_items = [
    "DDoS — Перегрузка канала. Сайт колледжа лежит.",
    "Man-in-the-Middle (MITM) — Прослушивание трафика.",
    "ARP-Spoofing / DNS-Spoofing — Подмена адресов.",
    "Брутфорс сервисов — Подбор паролей к RDP, SSH."
]
add_content_slide(prs, "ВЕКТОР 3. СЕТЬ И СЕРВИСЫ", network_items, NEON_YELLOW)

# Слайд 6: Zero-day
zeroday_items = [
    "Уязвимости, о которых ещё никто не знает",
    "❌ Нет патчей",
    "❌ Нет сигнатур",
    "❌ Нет защиты",
    "✅ Только мониторинг аномалий",
    "",
    "ЭТО НЕ БАГ, ЭТО ФИЧА?"
]
add_content_slide(prs, "ВЕКТОР 4. УГРОЗЫ НУЛЕВОГО ДНЯ (ZERO-DAY)", zeroday_items, NEON_BLUE)

# Слайд 7: Итог
summary_items = [
    "🔴 Люди → Обучай и контролируй",
    "🟠 Вредоносное ПО → Ограничивай запуск",
    "🟡 Сеть → Сегментируй и шифруй",
    "🔵 Zero-day → Мониторь аномалии"
]
add_content_slide(prs, "МОДЕЛЬ УГРОЗ СИСТЕМНОГО АДМИНИСТРАТОРА", summary_items, WHITE)

# Сохраняем презентацию
prs.save('Cyber_Threats_Presentation.pptx')
print("✅ Презентация успешно создана: Cyber_Threats_Presentation.pptx")
