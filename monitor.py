#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Мониторинг CPU и RAM для практической работы.

Примеры:
  py monitor.py --inventory-only
  py monitor.py --mode idle --samples 30 --interval 1
  py monitor.py --mode workload --samples 30 --interval 1

Сохраняет CSV, сведения об оборудовании в JSON, статистику и графики.
"""
from __future__ import annotations
import argparse
import csv
import json
import platform
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path
import psutil

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
GRAPHS = ROOT / "graphs"
DATA.mkdir(exist_ok=True)
GRAPHS.mkdir(exist_ok=True)
FIELDS = ["timestamp", "mode", "cpu_percent", "ram_used_gib", "ram_available_gib", "ram_percent"]


def run_powershell(script: str):
    if sys.platform != "win32":
        return None
    exe = shutil.which("powershell") or shutil.which("pwsh")
    if not exe:
        return None
    try:
        result = subprocess.run(
            [exe, "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=20
        )
        if result.returncode != 0 or not result.stdout.strip():
            return None
        raw = result.stdout.strip()
        start = raw.find("{")
        return json.loads(raw[start:]) if start >= 0 else None
    except Exception:
        return None


def get_windows_inventory():
    script = r'''
$cpu = Get-CimInstance Win32_Processor | Select-Object -First 1 Name,NumberOfCores,NumberOfLogicalProcessors,L3CacheSize
$board = Get-CimInstance Win32_BaseBoard | Select-Object -First 1 Manufacturer,Product,Version
$ram = @(Get-CimInstance Win32_PhysicalMemory | Select-Object Manufacturer,PartNumber,Capacity,Speed,ConfiguredClockSpeed,SMBIOSMemoryType)
$gpu = @(Get-CimInstance Win32_VideoController | Select-Object Name,AdapterRAM,DriverVersion,VideoProcessor)
$disks = @(Get-CimInstance Win32_DiskDrive | Select-Object Model,Size,InterfaceType,MediaType)
$physical = @(Get-PhysicalDisk -ErrorAction SilentlyContinue | Select-Object FriendlyName,MediaType,BusType,Size)
$bios = Get-CimInstance Win32_BIOS | Select-Object -First 1 SMBIOSBIOSVersion,ReleaseDate,Manufacturer
$system = Get-CimInstance Win32_ComputerSystem | Select-Object -First 1 Manufacturer,Model
[pscustomobject]@{cpu=$cpu; motherboard=$board; ram=$ram; gpu=$gpu; disks=$disks; physical_disks=$physical; bios=$bios; system=$system} | ConvertTo-Json -Depth 6 -Compress
'''
    return run_powershell(script)


def gib(value):
    try:
        return round(float(value) / (1024 ** 3), 2)
    except Exception:
        return None


def as_list(value):
    if value is None:
        return []
    return value if isinstance(value, list) else [value]


def detect_gpu_vram():
    nvsmi = shutil.which("nvidia-smi")
    if not nvsmi:
        return None
    try:
        r = subprocess.run(
            [nvsmi, "--query-gpu=name,memory.total", "--format=csv,noheader,nounits"],
            capture_output=True, text=True, timeout=10
        )
        if r.returncode != 0:
            return None
        rows = []
        for line in r.stdout.strip().splitlines():
            parts = [p.strip() for p in line.split(",")]
            if len(parts) >= 2:
                try:
                    rows.append({"model": parts[0], "vram_gib": round(float(parts[1]) / 1024, 2), "source": "nvidia-smi"})
                except ValueError:
                    pass
        return rows or None
    except Exception:
        return None


def detect_inventory():
    vm = psutil.virtual_memory()
    result = {
        "captured_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "computer_model": "Не определено",
        "os": {"system": platform.system(), "release": platform.release(), "version": platform.version(), "architecture": platform.machine()},
        "cpu": {"model": platform.processor() or "Не определено", "physical_cores": psutil.cpu_count(logical=False), "logical_processors": psutil.cpu_count(logical=True), "l3_cache_mb": None},
        "motherboard": None,
        "ram": {"total_gib": round(vm.total / (1024 ** 3), 2), "type": None, "modules": None, "speed_mhz": None, "details": []},
        "gpu": [], "disks": [], "bios": None, "psu": None, "cooling": None, "case": None,
        "limitations": [
            "Блок питания и модель воздушного/жидкостного охлаждения обычно не доступны программному опросу.",
            "Если nvidia-smi недоступен, объём VRAM из Win32_VideoController может быть ограничен способом представления Windows."
        ]
    }
    wmi = get_windows_inventory()
    if wmi:
        system = wmi.get("system") or {}
        result["computer_model"] = " ".join(filter(None, [system.get("Manufacturer"), system.get("Model")])) or result["computer_model"]
        cpu = wmi.get("cpu") or {}
        if cpu:
            result["cpu"] = {
                "model": cpu.get("Name") or result["cpu"]["model"],
                "physical_cores": cpu.get("NumberOfCores") or result["cpu"]["physical_cores"],
                "logical_processors": cpu.get("NumberOfLogicalProcessors") or result["cpu"]["logical_processors"],
                "l3_cache_mb": round(cpu["L3CacheSize"] / 1024, 2) if cpu.get("L3CacheSize") else None
            }
        result["motherboard"] = wmi.get("motherboard")
        result["bios"] = wmi.get("bios")
        modules = as_list(wmi.get("ram"))
        type_map = {24: "DDR3", 26: "DDR4", 34: "DDR5"}
        result["ram"]["modules"] = len(modules) if modules else None
        result["ram"]["details"] = [{
            "manufacturer": m.get("Manufacturer"), "part_number": m.get("PartNumber"),
            "capacity_gib": gib(m.get("Capacity")),
            "speed_mhz": m.get("ConfiguredClockSpeed") or m.get("Speed"),
            "type": type_map.get(m.get("SMBIOSMemoryType"), str(m.get("SMBIOSMemoryType")))
        } for m in modules]
        types = sorted(set(x["type"] for x in result["ram"]["details"] if x.get("type")))
        speeds = [x["speed_mhz"] for x in result["ram"]["details"] if x.get("speed_mhz")]
        result["ram"]["type"] = ", ".join(types) if types else None
        result["ram"]["speed_mhz"] = max(speeds) if speeds else None
        physical = as_list(wmi.get("physical_disks"))
        result["physical_disks"] = [{
            "model": p.get("FriendlyName"), "media_type": p.get("MediaType"),
            "bus_type": p.get("BusType"), "capacity_gib": gib(p.get("Size"))
        } for p in physical]
        physical_by_name = {p.get("FriendlyName"): p for p in physical if p.get("FriendlyName")}
        for d in as_list(wmi.get("disks")):
            p = physical_by_name.get(d.get("Model"), {})
            result["disks"].append({
                "model": d.get("Model"), "capacity_gib": gib(d.get("Size")),
                "interface": d.get("InterfaceType"), "media_type": d.get("MediaType"),
                "bus_type": p.get("BusType"), "physical_media": p.get("MediaType")
            })
        for g in as_list(wmi.get("gpu")):
            result["gpu"].append({
                "model": g.get("Name"), "vram_gib_windows_reported": gib(g.get("AdapterRAM")),
                "driver": g.get("DriverVersion"), "video_processor": g.get("VideoProcessor")
            })
    nv = detect_gpu_vram()
    if nv:
        result["gpu_nvidia_smi"] = nv
    return result


def load_rows(path):
    if not path.exists():
        return []
    with path.open("r", newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def aggregate(rows):
    out = {}
    for key in ("cpu_percent", "ram_used_gib", "ram_available_gib", "ram_percent"):
        values = [float(r[key]) for r in rows if r.get(key) not in (None, "")]
        if values:
            out[key] = {"min": round(min(values), 2), "avg": round(sum(values) / len(values), 2), "max": round(max(values), 2)}
    return out


def create_graphs():
    import matplotlib.pyplot as plt
    idle, work = load_rows(DATA / "idle.csv"), load_rows(DATA / "workload.csv")
    if not idle or not work:
        return False
    def values(rows, key):
        return [float(r[key]) for r in rows]
    fig, ax = plt.subplots(figsize=(10, 4.8))
    ax.plot(range(1, len(idle)+1), values(idle, "cpu_percent"), label="Простой")
    ax.plot(range(1, len(work)+1), values(work, "cpu_percent"), label="Рабочая нагрузка")
    ax.set_title("Загрузка CPU"); ax.set_xlabel("Номер измерения"); ax.set_ylabel("%")
    ax.grid(True, alpha=0.25); ax.legend(); fig.tight_layout()
    fig.savefig(GRAPHS / "cpu_comparison.png", dpi=160); plt.close(fig)
    fig, ax = plt.subplots(figsize=(10, 4.8))
    ax.plot(range(1, len(idle)+1), values(idle, "ram_used_gib"), label="Простой")
    ax.plot(range(1, len(work)+1), values(work, "ram_used_gib"), label="Рабочая нагрузка")
    ax.set_title("Использование оперативной памяти"); ax.set_xlabel("Номер измерения"); ax.set_ylabel("Занято, ГиБ")
    ax.grid(True, alpha=0.25); ax.legend(); fig.tight_layout()
    fig.savefig(GRAPHS / "ram_comparison.png", dpi=160); plt.close(fig)
    return True


def refresh_summary(interval_s=1.0):
    machine_path = DATA / "machine.json"
    machine = json.loads(machine_path.read_text(encoding="utf-8")) if machine_path.exists() else detect_inventory()
    runs = {}
    for mode in ("idle", "workload"):
        rows = load_rows(DATA / f"{mode}.csv")
        if rows:
            runs[mode] = {"count": len(rows), "interval_s": interval_s, "metrics": aggregate(rows)}
    comparison = {}
    if "idle" in runs and "workload" in runs:
        for metric in ("cpu_percent", "ram_used_gib", "ram_available_gib", "ram_percent"):
            comparison[metric] = {"idle": runs["idle"]["metrics"].get(metric), "workload": runs["workload"]["metrics"].get(metric)}
    data = {"status": "Оба измерения собраны" if len(runs) == 2 else "Собраны не все измерения",
            "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
            "machine": machine, "runs": runs, "comparison": comparison}
    (DATA / "summary.json").write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    if len(runs) == 2:
        create_graphs()
    return data


def run_measurement(mode, samples, interval):
    path = DATA / f"{mode}.csv"
    print(f"\nРежим: {mode}. Будет записано {samples} точек с интервалом около {interval} сек.")
    if mode == "idle":
        print("Закройте тяжёлые программы и оставьте компьютер в обычном простое.")
    else:
        print("Откройте IDE, браузер примерно с 10 вкладками и запустите локальную программу.")
    input("Когда система готова, нажмите Enter...")
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDS); writer.writeheader()
        for i in range(samples):
            cpu = psutil.cpu_percent(interval=interval)
            mem = psutil.virtual_memory()
            row = {"timestamp": datetime.now().astimezone().isoformat(timespec="seconds"), "mode": mode,
                   "cpu_percent": round(cpu, 2), "ram_used_gib": round(mem.used/(1024**3), 2),
                   "ram_available_gib": round(mem.available/(1024**3), 2), "ram_percent": round(mem.percent, 2)}
            writer.writerow(row); f.flush()
            print(f"{i+1:02d}/{samples}: CPU {row['cpu_percent']:5.1f}% | RAM {row['ram_used_gib']:5.2f} ГиБ занято | {row['ram_available_gib']:5.2f} ГиБ доступно")
    print(f"\nФайл сохранён: {path}")
    return refresh_summary()


def main():
    parser = argparse.ArgumentParser(description="Мониторинг CPU и RAM для практической работы.")
    parser.add_argument("--mode", choices=["idle", "workload"], help="Тип измерения")
    parser.add_argument("--samples", type=int, default=30, help="Количество точек; по заданию 30")
    parser.add_argument("--interval", type=float, default=1.0, help="Интервал в секундах")
    parser.add_argument("--inventory-only", action="store_true", help="Только собрать сведения об оборудовании")
    args = parser.parse_args()
    if args.inventory_only:
        inv = detect_inventory()
        (DATA / "machine.json").write_text(json.dumps(inv, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps(inv, ensure_ascii=False, indent=2)); return
    if not args.mode:
        parser.error("Укажите --mode idle или --mode workload.")
    if args.samples < 1 or args.interval <= 0:
        parser.error("samples и interval должны быть положительными.")
    inv = detect_inventory()
    (DATA / "machine.json").write_text(json.dumps(inv, ensure_ascii=False, indent=2), encoding="utf-8")
    run_measurement(args.mode, args.samples, args.interval)

if __name__ == "__main__":
    main()
