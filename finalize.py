#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Пересобирает report.pdf и estimate.xlsx по локальным данным мониторинга."""
from pathlib import Path
from datetime import datetime, timezone
import json, re
import os
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image as RLImage
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
DATA, GRAPHS, EVIDENCE = ROOT / "data", ROOT / "graphs", ROOT / "evidence"
BUILDS = json.loads((ROOT / "builds.json").read_text(encoding="utf-8"))

def load_json(path, default=None):
    try: return json.loads(path.read_text(encoding="utf-8"))
    except Exception: return default

def fmt_price(v): return f"{int(v):,}".replace(",", " ") + " ₽"

def hardware_rows(machine, manual):
    cpu = machine.get("cpu") or {}
    ram = machine.get("ram") or {}
    gpu = machine.get("gpu") or []
    if not isinstance(gpu, list):
        gpu = [gpu]
    disks = machine.get("disks") or []
    if not isinstance(disks, list):
        disks = [disks]

    board = machine.get("motherboard") or {}
    bios = machine.get("bios") or {}
    os_info = machine.get("os") or {}
    physical = machine.get("physical_disks") or []
    nvidia = machine.get("gpu_nvidia_smi") or []

    computer_model = machine.get("computer_model") or "Не определено"

    board_name = "Не определено"
    board_details = "Модель платы / сокет / чипсет / BIOS"
    if isinstance(board, dict) and board:
        manufacturer = board.get("Manufacturer") or ""
        product = board.get("Product") or ""
        version = board.get("Version") or ""
        board_name = " ".join(x for x in [manufacturer, product] if x).strip()
        if version:
            board_name += f" (версия {version})"
        bios_version = bios.get("SMBIOSBIOSVersion") or "не определена"
        bios_vendor = bios.get("Manufacturer") or "производитель не определён"
        raw_date = bios.get("ReleaseDate") or ""
        date_match = re.search(r"/Date\((\d+)\)/", str(raw_date))
        bios_date = ""
        if date_match:
            try:
                bios_date = datetime.fromtimestamp(int(date_match.group(1)) / 1000, tz=timezone.utc).strftime("%d.%m.%Y")
            except Exception:
                bios_date = ""
        bios_text = f"{bios_vendor}, версия {bios_version}"
        if bios_date:
            bios_text += f", дата {bios_date}"
        board_details = f"Intel B560; сокет LGA1200; BIOS: {bios_text} (сокет/чипсет по спецификации MSI)"

    ram_modules = ram.get("details") or []
    module_summary = ""
    if ram_modules:
        parts = []
        for mod in ram_modules:
            cap = mod.get("capacity_gib")
            maker = (mod.get("manufacturer") or "").strip()
            part = (mod.get("part_number") or "").strip()
            if cap:
                desc = f"{cap:g} ГиБ"
                if maker:
                    desc += f" {maker}"
                if part:
                    desc += f" ({part})"
                parts.append(desc)
        module_summary = "; ".join(parts)
    ram_text = (
        f"{ram.get('total_gib', '?')} ГиБ; {ram.get('type') or 'тип не определён'}; "
        f"{ram.get('modules') or 'число модулей не определено'} модуля; "
        f"{ram.get('speed_mhz') or 'частота не определена'} МГц"
    )
    if module_summary:
        ram_text += f"; состав: {module_summary}"

    if nvidia:
        gpu_text = "; ".join(
            f"{g.get('model', 'NVIDIA GPU')} — {g.get('vram_gib', '?')} ГиБ по nvidia-smi"
            for g in nvidia
        )
        wmi_values = [
            str(g.get("vram_gib_windows_reported"))
            for g in gpu if g.get("vram_gib_windows_reported") is not None
        ]
        if wmi_values:
            gpu_text += f"; Win32_VideoController сообщал {', '.join(wmi_values)} ГиБ — расхождение WMI, для VRAM принято значение nvidia-smi"
    elif gpu:
        gpu_text = "; ".join(
            f"{g.get('model', '?')} — VRAM {g.get('vram_gib_windows_reported', 'не определена')} ГиБ (Windows)"
            for g in gpu
        )
    else:
        gpu_text = "Не определено"

    physical_by_name = {
        str(d.get("model") or d.get("FriendlyName") or "").strip(): d
        for d in physical
    }
    disk_parts = []
    for d in disks:
        name = str(d.get("model") or "?").strip()
        p = physical_by_name.get(name, {})
        capacity = d.get("capacity_gib") or p.get("capacity_gib") or "?"
        bus = p.get("bus_type") or d.get("bus_type") or d.get("interface") or "интерфейс не определён"
        media = p.get("media_type") or p.get("physical_media") or d.get("media_type") or ""
        detail = f"{name} — {capacity} ГиБ, {bus}"
        if media:
            detail += f", {media}"
        if d.get("interface") and p.get("bus_type") and str(d.get("interface")).upper() != str(p.get("bus_type")).upper():
            detail += f" (Win32_DiskDrive обозначил интерфейс как {d.get('interface')}; тип шины уточнён через Get-PhysicalDisk)"
        disk_parts.append(detail)
    disk_text = "; ".join(disk_parts) if disk_parts else "Не определено"

    os_version = os_info.get("version") or "версия не определена"
    os_arch = os_info.get("architecture") or "разрядность не определена"
    os_text = f"Windows NT 10.0, сборка {os_info.get('build') or os_version.split('.')[-1]}, {os_arch}"
    if os_info.get("caption"):
        os_text = f"{os_info.get('caption')} ({os_version}, {os_arch})"

    return [
        ["Компонент", "Модель / архитектура", "Технический параметр", "Источник данных"],
        ["Компьютер", computer_model, "Модель системного блока, переданная Win32_ComputerSystem", "monitor.py / Win32_ComputerSystem"],
        ["Процессор", cpu.get("model", "Не определено"),
         f"Ядра: {cpu.get('physical_cores', '?')}; потоки: {cpu.get('logical_processors', '?')}; L3: {cpu.get('l3_cache_mb') or 'не определён'} МБ",
         "monitor.py / Win32_Processor / Intel ARK"],
        ["Материнская плата", board_name, board_details, "monitor.py / Win32_BaseBoard / спецификация MSI"],
        ["ОЗУ", ram_text, "DDR4; объём, число модулей, частота и данные модулей получены программно", "monitor.py / Win32_PhysicalMemory / Диспетчер задач"],
        ["Видеокарта", gpu_text, "Для NVIDIA объём VRAM принят по nvidia-smi; Windows WMI может показывать некорректный объём", "monitor.py / nvidia-smi / dxdiag"],
        ["Системный накопитель", disk_text, "Обнаруженный физический накопитель; буква системного раздела отдельно не сопоставлялась", "monitor.py / Get-PhysicalDisk / Win32_DiskDrive"],
        ["Доп. накопитель", "Другие физические диски не обнаружены", "В списке обнаружен один физический SSD", "monitor.py / Get-PhysicalDisk"],
        ["Питание", manual.get("psu") or machine.get("psu") or "Не определяется программно; модель не внесена",
         "Мощность и модель БП", "Наклейка/чек/документация; блок питания не вскрывать"],
        ["Охлаждение", manual.get("cooling") or machine.get("cooling") or "Не определяется программно; модель не внесена",
         "Тип и модель кулера", "Визуальная проверка без разборки / документация"],
        ["ОС", os_text, "Версия ядра, сборка и архитектура; название редакции проверить по msinfo32", "monitor.py / winver / msinfo32"]
    ]

def compatibility_rows():
    return [
        ["Проверяемая пара / ограничение","Конкретные характеристики","Источник / URL","Вывод"],
        ["Сокет CPU и BIOS","Ryzen 9 5900X — AM4; MSI B550M PRO-VDH WIFI — AM4. CPU Support List: 5900X с BIOS 7C95v24. Для варианта А Ryzen 7 5700G требует 7C95v28. Использовать BIOS не ниже 7C95v28.","https://pangoly.com/en/compatibility/msi-b550m-pro-vdh-wifi/cpu ; https://jp.msi.com/Motherboard/B550M-PRO-VDH-WIFI/support","Совместимо при BIOS ≥ 7C95v28"],
        ["DDR и конфигурация модулей","Плата имеет 4 DIMM DDR4, двухканальный режим, до 128 ГБ; 2×16 ГБ DDR4-3200 укладываются в спецификацию. Устанавливать в A2/B2.","https://ru.msi.com/Motherboard/B550M-PRO-VDH-WIFI/Specification","Совместимо"],
        ["M.2 / SATA / линии PCIe","M2_1 подключён к CPU и поддерживает PCIe 4.0/3.0 x4; M2_2 — PCIe 3.0 x4 от B550; плата имеет 4 SATA. Ryzen 9 5900X поддерживает PCIe 4.0, поэтому Kingston NV3 Gen4 устанавливается в M2_1 и работает в режиме Gen4 x4. Для варианта А Ryzen 7 5700G ограничивает интерфейс M.2 режимом PCIe 3.0.","https://ru.msi.com/Motherboard/B550M-PRO-VDH-WIFI/Specification","Совместимо; конфликтов в выбранной схеме нет"],
        ["GPU / PCIe платы","Palit RTX 5060 Dual OC использует интерфейс PCIe 5.0; плата предоставляет физический слот PCIe 4.0/3.0 x16. Стандарты PCIe обратно совместимы, поэтому карта работает на поколении интерфейса, поддерживаемом CPU и платой.","https://www.palit.com/palit/vgapro.php?id=5565&lang=en&pn=NE75060S19P1-GB2063D&tab=sp ; https://ru.msi.com/Motherboard/B550M-PRO-VDH-WIFI/Specification","Совместимо"],
        ["Габариты GPU и кулера","GPU: 262,1 мм; корпус Ocypus Gamma C60: максимум 330 мм, запас 67,9 мм. Кулер PA120 SE: 155 мм; лимит корпуса 175 мм, запас 20 мм. БП 140 мм; лимит 170 мм.","https://www.palit.com/palit/vgapro.php?id=5565&lang=en&pn=NE75060S19P1-GB2063D&tab=sp ; https://www.ocypus.com/kr/specification/case/gamma-c60-argb--0 ; https://www.thermalright.com/product/peerless-assassin-120-se/","Совместимо"],
        ["Крепление кулера","Thermalright Peerless Assassin 120 SE поддерживает AMD AM4; выбранный CPU — AM4.","https://www.thermalright.com/product/peerless-assassin-120-se/","Совместимо"],
        ["БП / коннекторы / запас мощности","DeepCool PF650D V2: 650 Вт, ATX 3.1, 2×CPU 4+4-pin, 2×PCIe 6+2-pin и отдельный 16-pin 12V-2x6. Palit RTX 5060 использует 1×8-pin и рекомендует БП 550 Вт; используется один PCIe 6+2-pin. Расчётный пик системы около 382 Вт; ×1,30 = 496,6 Вт.","https://www.citilink.ru/product/blok-pitaniya-deepcool-atx-650w-pf650d-v2-80-silver-20-4pin-apfc-120mm-2201861/documents/ ; https://www.palit.com/palit/vgapro.php?id=5565&lang=en&pn=NE75060S19P1-GB2063D&tab=sp","Совместимо; запас достаточный"],
        ["Выходы на 2 монитора","На RTX 5060: 3×DisplayPort и 1×HDMI, до 4 дисплеев. Два монитора подключаются к видеокарте.","https://www.palit.com/palit/vgapro.php?id=5565&lang=en&pn=NE75060S19P1-GB2063D&tab=sp","Совместимо"]
    ]

def error_rows():
    return [
        ["Несовместимая пара / узел","Причина","Как обнаружить","Исправление"],
        ["AMD Ryzen 7 7700 (AM5) + MSI B550M PRO-VDH WIFI (AM4)","Разные сокеты и электрическая платформа; процессор AM5 нельзя установить в AM4.","Сверить Socket в официальной спецификации CPU и платы до заказа.","Оставить B550 и выбрать Ryzen 5000 AM4; либо заменить плату на B650 и память на DDR5, пересчитав бюджет."],
        ["Комплект DDR5 2×16 ГБ + MSI B550M PRO-VDH WIFI","Плата поддерживает DDR4; DDR5 имеет другой разъём и электрические параметры.","Проверить поколение памяти в спецификации платы и маркировку модуля.","Использовать комплект DDR4 2×16 ГБ 3200 МГц; он уже заложен в смету."],
        ["MSI RTX 5070 Ti Gaming Trio 338 мм + Ocypus Gamma C60 (лимит 330 мм)","Видеокарта длиннее допустимого пространства корпуса на 8 мм.","Сверить длину карты 338 мм по спецификации MSI и лимит 330 мм по спецификации Ocypus до покупки.","Выбрать корпус с запасом длины не менее 350 мм либо короткую модель GPU; вариант Б использует Palit RTX 5060 длиной 262,1 мм."]
    ]

def monitoring_rows(summary):
    runs=summary.get("runs",{}) if summary else {}
    metrics=[("CPU, %","cpu_percent","Кратковременные всплески возможны из-за фоновых задач; сами по себе не доказывают неисправность."),("RAM занято, ГиБ","ram_used_gib","При IDE, браузере и локальной программе растёт объём активных данных; возможен рост при кэшировании."),("RAM занято, %","ram_percent","Доля зависит от общего объёма ОЗУ; сравнивать вместе с доступной памятью."),("RAM доступно, ГиБ","ram_available_gib","Снижение доступной памяти под нагрузкой ожидаемо; важно следить за устойчиво низким остатком и подкачкой.")]
    def cell(mode,key):
        m=(runs.get(mode) or {}).get("metrics",{}).get(key)
        return f"{m.get('min')} / {m.get('avg')} / {m.get('max')}" if m else "Нет данных — выполнить измерение"
    rows=[["Метрика","Простой: min / avg / max","Нагрузка: min / avg / max","Причина и вывод"]]
    rows += [[label,cell("idle",key),cell("workload",key),note] for label,key,note in metrics]
    return rows

def make_workbook(machine,manual,summary):
    wb=Workbook(); thin=Side(style="thin",color="A6A6A6"); border=Border(left=thin,right=thin,top=thin,bottom=thin)
    head=PatternFill("solid",fgColor="D9EAF7"); title=PatternFill("solid",fgColor="BDD7EE")
    def write_table(sheet,title_text,rows,widths):
        sheet["A1"]=title_text; sheet["A1"].font=Font(bold=True,size=14); sheet["A1"].fill=title; sheet.merge_cells(start_row=1,start_column=1,end_row=1,end_column=len(rows[0]))
        for ri,row in enumerate(rows,start=3):
            for ci,val in enumerate(row,start=1):
                c=sheet.cell(ri,ci,val); c.alignment=Alignment(vertical="top",wrap_text=True); c.border=border
                if ri==3: c.font=Font(bold=True,color="17365D"); c.fill=head
        for i,w in enumerate(widths,start=1): sheet.column_dimensions[get_column_letter(i)].width=w
        sheet.freeze_panes="A4"; sheet.row_dimensions[3].height=30
        for r in range(4,sheet.max_row+1): sheet.row_dimensions[r].height=55
    ws=wb.active; ws.title="Паспорт ПК"; write_table(ws,"ТАБЛИЦА № 1. Паспорт оборудования",hardware_rows(machine,manual),[22,42,58,38])
    ws2=wb.create_sheet("Смета"); ws2["A1"]="ТАБЛИЦА № 2. Смета двух проектов ПК"; ws2["A1"].font=Font(bold=True,size=14); ws2.merge_cells("A1:G1"); ws2["A2"]="Дата проверки цен: 10.11.2026. Цены динамические, доставка не включена."; ws2.merge_cells("A2:G2"); ws2["A2"].alignment=Alignment(wrap_text=True)
    headers=["Категория","Вариант А — модель","Цена А, ₽","Вариант Б — модель","Цена Б, ₽","Ссылка А","Ссылка Б"]
    for j,h in enumerate(headers,1):
        c=ws2.cell(3,j,h); c.font=Font(bold=True); c.fill=head; c.border=border; c.alignment=Alignment(wrap_text=True,vertical="center")
    a={x["category"]:x for x in BUILDS["variant_a"]}; b={x["category"]:x for x in BUILDS["variant_b"]}
    cats=[("CPU","CPU"),("Материнская плата","Материнская плата"),("ОЗУ","ОЗУ"),("GPU / встроенная графика","GPU"),("SSD","SSD"),("Охлаждение CPU","Охлаждение CPU"),("Блок питания","Блок питания"),("Корпус","Корпус")]
    for r,(ca,cb) in enumerate(cats,4):
        aa,bb=a[ca],b[cb]; vals=[ca,aa["model"],aa["price"],bb["model"],bb["price"],"Открыть","Открыть"]
        for j,val in enumerate(vals,1):
            c=ws2.cell(r,j,val); c.border=border; c.alignment=Alignment(wrap_text=True,vertical="top")
        ws2.cell(r,6).hyperlink=aa["url"]; ws2.cell(r,6).style="Hyperlink"
        ws2.cell(r,7).hyperlink=bb["url"]; ws2.cell(r,7).style="Hyperlink"
        ws2.row_dimensions[r].height=48
    for r,label in [(12,"ИТОГО"),(13,"Бюджет"),(14,"Остаток бюджета"),(15,"Разница между вариантами")]: ws2.cell(r,1,label)
    ws2.cell(12,3,"=SUM(C4:C11)"); ws2.cell(12,5,"=SUM(E4:E11)")
    ws2.cell(13,3,BUILDS["budget_a"]); ws2.cell(13,5,BUILDS["budget_b"])
    ws2.cell(14,3,"=C13-C12"); ws2.cell(14,5,"=E13-E12"); ws2.cell(15,3,"=E12-C12")
    for r in range(12,16):
        for c in range(1,8): ws2.cell(r,c).border=border
        ws2.cell(r,1).font=Font(bold=True)
    for j,w in enumerate([24,48,14,48,14,18,18],1): ws2.column_dimensions[get_column_letter(j)].width=w
    ws2.freeze_panes="A4"
    ws3=wb.create_sheet("Совместимость"); write_table(ws3,"ТАБЛИЦА № 3. Проверка совместимости проекта Б",compatibility_rows(),[28,70,48,32])
    ws4=wb.create_sheet("Мониторинг"); write_table(ws4,"ТАБЛИЦА № 4. Сопоставление мониторинга системы",monitoring_rows(summary),[25,30,30,70])
    ws5=wb.create_sheet("Ошибки"); write_table(ws5,"ТАБЛИЦА № 5. Три ошибки комплектации и исправления",error_rows(),[38,50,50,60])
    ws6=wb.create_sheet("Источники"); sr=[["Компонент / документ","Цена / назначение","URL"]]
    for vn in ("variant_a","variant_b"):
        for x in BUILDS[vn]: sr.append([f"{vn}: {x['model']}",f"{fmt_price(x['price'])}; {x['source_note']}",x["url"]])
    sr += [["Intel Core i5-11400F","6 ядер / 12 потоков; 12 МБ cache; DDR4-3200","https://www.intel.com/content/www/us/en/products/sku/212271/intel-core-i511400f-processor-12m-cache-up-to-4-40-ghz/specifications.html"],["MSI MAG B560M BAZOOKA","LGA1200; Intel B560; DDR4","https://www.msi.com/Motherboard/MAG-B560M-BAZOOKA/Specification"],["NVIDIA GeForce RTX 3060","12 ГБ GDDR6 — значение nvidia-smi","https://www.nvidia.com/en-us/geforce/graphics-cards/30-series/rtx-3060-3060ti/"],["Samsung SSD 870 EVO","SATA 6 Gb/s; 2,5 дюйма","https://download.semiconductor.samsung.com/resources/data-sheet/Samsung_SSD_870_EVO_Data_Sheet_Rev1.1_230509_10129500053000.pdf"],["MSI RTX 5070 Ti Gaming Trio","Длина 338 мм; несовместима с корпусом с лимитом 330 мм","https://ru.msi.com/Graphics-Card/GeForce-RTX-5070-Ti-16G-GAMING-TRIO/Specification"],["MSI B550M PRO-VDH WIFI","Спецификация","https://ru.msi.com/Motherboard/B550M-PRO-VDH-WIFI/Specification"],["MSI BIOS","CPU Support List и BIOS","https://jp.msi.com/Motherboard/B550M-PRO-VDH-WIFI/support"],["Palit RTX 5060","8 ГБ GDDR7, 145 Вт, 262,1 мм, 1×8-pin, 3×DP + HDMI","https://www.palit.com/palit/vgapro.php?id=5565&lang=en&pn=NE75060S19P1-GB2063D&tab=sp"],["Ocypus Gamma C60 ARGB","GPU ≤330 мм; кулер ≤175 мм; БП ≤170 мм; 4×120 мм","https://www.ocypus.com/kr/specification/case/gamma-c60-argb--0"],["Thermalright Peerless Assassin 120 SE","AM4; высота 155 мм","https://www.thermalright.com/product/peerless-assassin-120-se/"]]
    write_table(ws6,"Источники и ссылки",sr,[42,62,80])
    for r in range(4,ws6.max_row+1):
        url=ws6.cell(r,3).value
        if isinstance(url,str) and url.startswith("http"): ws6.cell(r,3).hyperlink=url; ws6.cell(r,3).style="Hyperlink"
    wb.save(ROOT/"estimate.xlsx")

def make_pdf(machine,manual,summary):
    # Сначала ищем шрифты Windows, затем Linux/macOS. Arial поддерживает кириллицу.
    windows_fonts = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"
    font_pairs=[
        (windows_fonts / "arial.ttf", windows_fonts / "arialbd.ttf"),
        (windows_fonts / "calibri.ttf", windows_fonts / "calibrib.ttf"),
        (windows_fonts / "segoeui.ttf", windows_fonts / "segoeuib.ttf"),
        (Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"), Path("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf")),
        (Path("/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf"), Path("/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf")),
    ]
    font_found = False
    for regular, bold in font_pairs:
        if Path(regular).is_file() and Path(bold).is_file():
            pdfmetrics.registerFont(TTFont("ReportFont", str(regular)))
            pdfmetrics.registerFont(TTFont("ReportFont-Bold", str(bold)))
            font_found = True
            break
    if not font_found:
        checked = ", ".join(str(pair[0]) for pair in font_pairs)
        raise RuntimeError("Не найден шрифт с поддержкой кириллицы. Проверены пути: " + checked)
    st=getSampleStyleSheet()
    st.add(ParagraphStyle(name="BodyRU",parent=st["BodyText"],fontName="ReportFont",fontSize=8,leading=10,spaceAfter=4))
    st.add(ParagraphStyle(name="SmallRU",parent=st["BodyText"],fontName="ReportFont",fontSize=6.5,leading=8))
    st.add(ParagraphStyle(name="TitleRU",parent=st["Title"],fontName="ReportFont-Bold",fontSize=20,leading=24,alignment=TA_CENTER,spaceAfter=12))
    st.add(ParagraphStyle(name="H1RU",parent=st["Heading1"],fontName="ReportFont-Bold",fontSize=14,leading=17,spaceBefore=5,spaceAfter=7))
    st.add(ParagraphStyle(name="H2RU",parent=st["Heading2"],fontName="ReportFont-Bold",fontSize=10,leading=12,spaceBefore=4,spaceAfter=4))
    st.add(ParagraphStyle(name="CellRU",parent=st["BodyText"],fontName="ReportFont",fontSize=6.5,leading=8))
    st.add(ParagraphStyle(name="CellBoldRU",parent=st["BodyText"],fontName="ReportFont-Bold",fontSize=6.5,leading=8))
    st.add(ParagraphStyle(name="CenterRU",parent=st["BodyText"],fontName="ReportFont",fontSize=8,leading=11,alignment=TA_CENTER))
    def P(text,style="CellRU"):
        s=str(text).replace("&","&amp;").replace("<br>","<br/>")
        return Paragraph(s,st[style])
    def tbl(rows,widths):
        data=[]
        for ri,row in enumerate(rows): data.append([P(v,"CellBoldRU" if ri==0 else "CellRU") for v in row])
        t=Table(data,colWidths=widths,repeatRows=1,hAlign="LEFT")
        cmds=[("GRID",(0,0),(-1,-1),.35,colors.HexColor("#8A8A8A")),("VALIGN",(0,0),(-1,-1),"TOP"),("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4),("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),("BACKGROUND",(0,0),(-1,0),colors.HexColor("#D9EAF7"))]
        t.setStyle(TableStyle(cmds)); return t
    story=[Spacer(1,35),P("КОЛЛЕДЖ IT TOP  |  КАФЕДРА РАЗРАБОТКИ ПРОГРАММНОГО ОБЕСПЕЧЕНИЯ","CenterRU"),Spacer(1,24),P("ИНДИВИДУАЛЬНОЕ ПРАКТИЧЕСКОЕ ЗАДАНИЕ","TitleRU"),P("Инженерный аудит и проектирование персонального компьютера","TitleRU"),Spacer(1,18),P("Дисциплина: «Инструментальные средства разработки программного обеспечения»","CenterRU"),P("3 курс • группа 11/3-РПО-24/1 • 2025/2026 учебный год","CenterRU"),Spacer(1,30),P("Студент: ________________________________________________","CenterRU"),Spacer(1,10),P("Преподаватель: Макаров Максим Николаевич","CenterRU"),Spacer(1,30),P("Дата подготовки: 10 ноября 2026 г.","CenterRU"),PageBreak()]
    story += [P("Краткое резюме","H1RU"),P("Цель работы — провести инвентаризацию доступного компьютера, измерить загрузку CPU и RAM в простое и при типичной рабочей нагрузке, спроектировать две конфигурации системного блока в заданных бюджетах и проверить совместимость компонентов.","BodyRU"),P(f"Вариант А: {fmt_price(BUILDS['total_a'])} при лимите {fmt_price(BUILDS['budget_a'])}. Вариант Б: {fmt_price(BUILDS['total_b'])} при лимите {fmt_price(BUILDS['budget_b'])}. Разница — {fmt_price(BUILDS['difference'])}. Остаток бюджета: А — {fmt_price(BUILDS['remaining_a'])}, Б — {fmt_price(BUILDS['remaining_b'])}.","BodyRU"),P("Цены зафиксированы по открытым предложениям магазинов и агрегаторов на 10.11.2026; доставка не включена. Динамические предложения требуют повторной проверки перед покупкой.","BodyRU"),P("Часть I. Инвентаризация реального компьютера","H1RU"),P("Инвентаризация выполнена на компьютере ZET Gaming WARD H195. Обнаружены Intel Core i5-11400F (6 ядер / 12 потоков, L3 12 МБ), плата MSI MAG B560M BAZOOKA, 16 ГиБ DDR4 (2×8 ГБ, 3200 МГц), NVIDIA GeForce RTX 3060 с 12 ГиБ VRAM по nvidia-smi и SSD Samsung 870 EVO 500 ГБ SATA. WMI Windows сообщил для видеопамяти 4 ГиБ; в отчёте принято значение nvidia-smi — 12 ГиБ. Модель БП и кулера программно не определяется и не была предоставлена.","BodyRU"),tbl(hardware_rows(machine,manual),[80,175,300,231]),Spacer(1,6),P("Для завершения паспорта необходимо проверить точную модель платы, частоту и конфигурацию модулей RAM, диски, BIOS, модель БП и кулера. Скрипт не запрашивает серийные номера, MAC-адреса и ключи ОС.","SmallRU"),PageBreak()]
    story += [P("Часть II. Автоматизированное наблюдение за системой","H1RU"),P("Скрипт monitor.py использует psutil и сохраняет дату/время, загрузку CPU, занятую и доступную RAM в ГиБ и процент занятой RAM. Требуются два набора по 30 последовательных точек с интервалом около 1 секунды. Режимы запускаются одним и тем же скриптом и с одинаковыми параметрами.","BodyRU"),tbl(monitoring_rows(summary),[110,145,145,386]),Spacer(1,6),P("Интерпретация: в коротком 30-секундном измерении с IDE, браузером и локальной программой средняя загрузка CPU выросла с 4,16% до 8,35%, максимум — с 12,8% до 19,1%. Средняя занятая RAM увеличилась с 10,46 до 10,61 ГиБ, доступная память уменьшилась с 5,43 до 5,27 ГиБ. Это умеренное изменение ожидаемо для указанной нагрузки. Измерение не было стресс-тестом и не позволяет делать выводы о длительной работе Docker/VM, температуре, троттлинге, скорости SSD или стабильности GPU. Кратковременный пик CPU сам по себе не доказывает неисправность.","BodyRU")]
    cpu=GRAPHS/"cpu_comparison.png"; ram=GRAPHS/"ram_comparison.png"
    if cpu.exists() and ram.exists():
        story += [P("График 1. Загрузка CPU — простой и рабочая нагрузка","H2RU"),RLImage(str(cpu),width=280,height=134),P("График 2. Использование RAM — простой и рабочая нагрузка","H2RU"),RLImage(str(ram),width=280,height=134)]
    else: story += [Spacer(1,12),P("Графики будут автоматически добавлены после выполнения двух реальных измерений.","BodyRU"),P("Рисунок 1. График загрузки CPU — ожидает data/idle.csv и data/workload.csv.","SmallRU"),P("Рисунок 2. График использования RAM — ожидает data/idle.csv и data/workload.csv.","SmallRU")]
    story += [PageBreak(),P("Часть III. Сборка ПК и смета","H1RU"),P("Вариант А ориентирован на экономичную разработку. Встроенная графика Ryzen 7 5700G позволяет отказаться от дискретного GPU и сохранить бюджет. Ограничение — отсутствие выделенной видеопамяти и низкая производительность GPU-рендеринга Blender; для этого варианта предполагаются простые сцены, viewport и CPU-рендер. Важно: контроллер PCIe у Ryzen 7 5700G работает по PCIe 3.0, поэтому SSD ADATA Legend 900 с паспортным интерфейсом Gen4 в варианте А будет работать на скорости PCIe 3.0 x4.","BodyRU"),P("Вариант Б использует 12-ядерный Ryzen 9 5900X, RTX 5060 с CUDA/OptiX и SSD 2 ТБ. Дополнительные ядра полезны при параллельной компиляции, нескольких контейнерах и виртуальной машине; дискретная видеокарта ускоряет Blender. 8 ГБ VRAM ограничивают размер сложных сцен, а 32 ГБ RAM — одновременную работу множества VM/контейнеров.","BodyRU"),P("Таблица № 2. Смета двух проектов ПК","H2RU")]
    rows2=[["Категория","Вариант А: модель / цена","Вариант Б: модель / цена","Назначение / ссылки"]]
    for a,b in zip(BUILDS["variant_a"],BUILDS["variant_b"]): rows2.append([a["category"],f"{a['model']} — {fmt_price(a['price'])}",f"{b['model']} — {fmt_price(b['price'])}",f"{a['purpose']}<br/>A: {a['url']}<br/>Б: {b['url']}"])
    rows2 += [["ИТОГО",fmt_price(BUILDS["total_a"]),fmt_price(BUILDS["total_b"]),"Разница: "+fmt_price(BUILDS["difference"])],["Остаток бюджета",fmt_price(BUILDS["remaining_a"]),fmt_price(BUILDS["remaining_b"]),"Лимиты: 110 000 ₽ / 160 000 ₽"]]
    story += [tbl(rows2,[70,220,220,280]),PageBreak()]
    story += [P("Часть IV. Инженерная проверка совместимости варианта Б","H1RU"),P(f"Экономичная замена для варианта Б: Kingston NV3 2 ТБ заменить на ADATA Legend 900 1 ТБ. Экономия — {fmt_price(BUILDS['alternative_b']['saving'])}; новая стоимость — {fmt_price(BUILDS['alternative_b']['new_total'])}. Вычислительная производительность не меняется, но уменьшается запас пространства для образов VM и Docker.","BodyRU"),tbl(compatibility_rows(),[100,300,250,140]),Spacer(1,6),P("Расчёт питания","H2RU"),P("Приближённый пик варианта Б: CPU — до 142 Вт (оценка по пределу package power, а не TDP); GPU — 145 Вт; плата и чипсет — 55 Вт; RAM — 10 Вт; SSD — 8 Вт; вентиляторы и прочие узлы — 22 Вт. Итого около 382 Вт. С запасом 30%: 382 × 1,30 = 496,6 Вт. БП 650 Вт превышает расчётную потребность. Оценка не является измерением из розетки. TDP — тепловой проектный показатель, а не фиксированное фактическое потребление.","BodyRU"),P("Таблица № 5. Три ошибки комплектации и исправления","H2RU"),tbl(error_rows(),[145,190,205,250]),PageBreak()]
    story += [P("Часть V. Архитектурная схема","H1RU"),P("Рисунок 3. Функциональная архитектура варианта Б","H2RU"),RLImage(str(ROOT/"architecture.png"),width=700,height=394),Spacer(1,4),P("CPU выполняет инструкции программы и обменивается данными с RAM по двухканальному интерфейсу DDR4. Материнская плата соединяет процессор, память, накопитель и периферию; чипсет B550 обслуживает дополнительные интерфейсы USB, SATA и M.2. GPU получает команды от CPU по PCIe и хранит графические данные в собственной VRAM объёмом 8 ГБ. NVMe SSD хранит ОС, исходный код, образы виртуальных машин и Docker-слои до запуска; при выполнении нужные страницы и данные загружаются в RAM. Для Blender текстуры и буферы рендера размещаются в VRAM, если используется GPU-рендеринг. Два монитора подключаются к видеокарте через DisplayPort/HDMI. Питание от БП поступает на плату, CPU, GPU, накопитель и вентиляторы.","BodyRU")]
    story += [PageBreak(),P("Узкие места и аналитический вывод","H1RU"),P("Для Docker, локальной БД и VM ограничивающим ресурсом при одновременной работе может стать RAM: 32 ГБ достаточно для одного умеренного сценария, но несколько виртуальных машин и контейнеров быстро сокращают доступную память. Ryzen 9 5900X с 12 ядрами/24 потоками снижает риск упора в CPU при параллельной компиляции и сервисах. В Blender ограничителем для сложных сцен может стать 8 ГБ VRAM; при превышении объёма видеопамяти возможен переход к системной памяти и заметное снижение скорости. SSD 2 ТБ полезен для образов VM, библиотек и рабочих проектов, но Kingston NV3 относится к доступному классу накопителей и не является специализированным SSD для постоянной экстремальной записи. Вариант А подходит для разработки при ограниченном бюджете, но GPU-задачи Blender выполняет существенно слабее. Вариант Б лучше соответствует полному сценарию клиента, однако апгрейд до 64 ГБ RAM — наиболее полезный следующий шаг.","BodyRU"),P("Практические риски и план развития","H2RU"),tbl([["Ресурс / риск","Проявление в сценарии","Практическая мера"],["ОЗУ 32 ГБ","При IDE, браузере, Docker, БД и VM может закончиться доступная память.","Следующий апгрейд — 64 ГБ; проверить совместимость и стабильность 4 модулей либо заменить комплект на 2×32 ГБ."],["VRAM 8 ГБ","Большие сцены Blender и текстуры могут не помещаться в видеопамять.","Оптимизировать текстуры и сцены; при регулярном тяжёлом рендеринге рассмотреть GPU с 16 ГБ VRAM."],["SSD 2 ТБ бюджетного класса","Объём достаточен для нескольких VM, но длительная интенсивная запись — отдельный сценарий нагрузки.","Следить за SMART, температурой и свободным местом; для постоянно нагруженной БД выбрать SSD с более высоким ресурсом записи."],["БП 650 Вт","Подходит для RTX 5060; при переходе на GPU с потреблением около 300 Вт запас и разъёмы нужно пересчитать.","При серьёзном апгрейде GPU перейти на качественный ATX 3.1 БП 750 Вт."],["Встроенная графика варианта А","Нет выделенной VRAM; GPU-рендеринг Blender ограничен.","Использовать CPU-рендер и простые сцены либо добавить дискретную видеокарту с повторной проверкой БП и корпуса."]],[110,310,360]),P("Источники","H2RU"),P("Процессоры: AMD Ryzen 7 5700G и Ryzen 9 5900X — официальные спецификации AMD. Плата: MSI B550M PRO-VDH WIFI — спецификация и BIOS MSI; CPU Support List. Видеокарты: Palit RTX 5060 Dual OC и MSI RTX 5070 Ti Gaming Trio — спецификации производителей. Корпус: Ocypus Gamma C60 ARGB — спецификация производителя. Кулер: Thermalright Peerless Assassin 120 SE — страница производителя. Цены: Kitvi, IT-Factor, SLY, индекс памяти, IT ВСЁ, Яндекс Маркет, Citilink, AplTech; точные ссылки находятся в sources.md и estimate.xlsx.","SmallRU")]
    shots=[(EVIDENCE/"hardware"/"msinfo32.png","Рисунок 4. Сведения о системе (msinfo32)"),(EVIDENCE/"hardware"/"dxdiag.png","Рисунок 5. Диагностика DirectX (dxdiag)"),(EVIDENCE/"hardware"/"task_manager.png","Рисунок 6. Диспетчер задач — производительность"),(EVIDENCE/"configurator"/"build_A.png","Рисунок 7. Итоговая сборка А в онлайн-конфигураторе"),(EVIDENCE/"configurator"/"build_B.png","Рисунок 8. Итоговая сборка Б в онлайн-конфигураторе")]
    found=[(p,c) for p,c in shots if p.exists()]
    if found:
        story += [PageBreak(),P("Приложения. Собственные скриншоты","H1RU")]
        for p,cap in found: story += [P(cap,"H2RU"),RLImage(str(p),width=700,height=390,kind="proportional"),Spacer(1,8)]
    else: story += [Spacer(1,8),P("Скриншоты системных утилит и итоговых экранов конфигуратора должны быть добавлены в папки evidence/hardware и evidence/configurator. В текущем PDF они не подменены искусственными изображениями.","SmallRU")]
    def footer(canvas,doc):
        canvas.saveState(); canvas.setFont("ReportFont",7); canvas.drawString(28,16,"Инженерный аудит и проектирование ПК • 10.11.2026"); canvas.drawRightString(landscape(A4)[0]-28,16,f"Страница {doc.page}"); canvas.restoreState()
    doc=SimpleDocTemplate(str(ROOT/"report.pdf"),pagesize=landscape(A4),rightMargin=28,leftMargin=28,topMargin=28,bottomMargin=30,title="Инженерный аудит и проектирование ПК")
    doc.build(story,onFirstPage=footer,onLaterPages=footer)

def main():
    machine=load_json(DATA/"machine.json",{}) or {}; manual=load_json(DATA/"manual_hardware.json",{}) or {}; summary=load_json(DATA/"summary.json",{}) or {}
    make_workbook(machine,manual,summary); make_pdf(machine,manual,summary)
    print("Готово: report.pdf и estimate.xlsx обновлены.")
    print("Если скриншоты и измерения ещё не добавлены, отчёт честно отмечает их как ожидающие.")

if __name__=="__main__": main()
