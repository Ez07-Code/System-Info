# InfoSystem_backend.py

import wmi
import socket
import os
import platform
import re

# --- Todas tus funciones de obtención de datos van aquí ---
# (get_system_info, get_bios_info, get_os_info, etc.)

def get_system_info(c):
    try:
        system = c.Win32_ComputerSystem()[0]
        return {"Fabricante": system.Manufacturer, "Modelo": system.Model}
    except Exception as e: return {"Error": f"No se pudo obtener la info del sistema: {e}"}

def get_bios_info(c):
    try:
        bios = c.Win32_BIOS()[0]
        return {"Fabricante BIOS": bios.Manufacturer, "Versión BIOS": bios.Version, "Número de Serie PC": bios.SerialNumber}
    except Exception as e: return {"Error": f"No se pudo obtener la info de la BIOS: {e}"}

def get_os_info(c):
    try:
        os_info = c.Win32_OperatingSystem()[0]
        return {"Sistema Operativo": os_info.Caption, "Versión": f"{platform.release()} ({os_info.Version})", "Arquitectura": os_info.OSArchitecture}
    except Exception as e: return {"Error": f"No se pudo obtener la info del SO: {e}"}

def get_network_info(c):
    hostname = socket.gethostname()
    ipv4_ethernet, ipv4_wifi = "No disponible o desconectado", "No disponible o desconectado"
    try:
        for adapter in c.Win32_NetworkAdapterConfiguration(IPEnabled=True):
            desc = adapter.Description.lower()
            if adapter.IPAddress:
                ip = adapter.IPAddress[0]
                if 'ethernet' in desc or 'gigabit' in desc: ipv4_ethernet = ip
                elif 'wi-fi' in desc or 'wireless' in desc: ipv4_wifi = ip
    except Exception: pass
    try:
        system = c.Win32_ComputerSystem()[0]
        return {"Hostname": hostname, "Usuario Actual": os.getlogin(), "Dominio/Grupo": system.Domain if system.PartOfDomain else system.Workgroup, "Dirección IPv4 (Ethernet)": ipv4_ethernet, "Dirección IPv4 (Wi-Fi)": ipv4_wifi}
    except Exception as e: return {"Error": f"No se pudo obtener la info de red: {e}"}

def get_ram_info(c):
    try:
        cs = c.Win32_ComputerSystem()[0]
        total_ram_gb = round(int(cs.TotalPhysicalMemory) / (1024**3), 2)
        mem_speed = "No disponible"
        physical_memory = c.Win32_PhysicalMemory()
        if physical_memory: mem_speed = f"{physical_memory[0].Speed} MHz"
        return {"RAM Total": f"{total_ram_gb} GB", "Velocidad": mem_speed}
    except Exception as e: return {"Error": f"No se pudo obtener la info de la RAM: {e}"}

def get_disk_info():
    disks = []
    try:
        c_storage = wmi.WMI(namespace="ROOT\Microsoft\Windows\Storage")
        physical_disks = c_storage.MSFT_PhysicalDisk()
        for disk in physical_disks:
            disks.append({"Fabricante": disk.Manufacturer, "Capacidad Total": f"{round(int(disk.Size) / (1024**3), 2)} GB", "Número de Serie": disk.SerialNumber.strip()})
        return disks
    except Exception:
        try:
            c_disk = wmi.WMI()
            for disk in c_disk.Win32_DiskDrive():
                disks.append({"Fabricante": disk.Model, "Capacidad Total": f"{round(int(disk.Size) / (1024**3), 2)} GB", "Número de Serie": disk.SerialNumber.strip() if disk.SerialNumber else "No disponible"})
            return disks
        except Exception as e_fallback:
             return [{"Error": f"No se pudo obtener la info de los discos: {e_fallback}"}]

def get_cpu_info(c):
    try:
        processor = c.Win32_Processor()[0]
        name = processor.Name.strip()
        manufacturer = processor.Manufacturer
        generation = "No detectada"
        match = re.search(r'i[3579]-(\d{1,2})\d{3}', name)
        if match: generation = f"{match.group(1)}ª Generación"
        elif "AMD" in manufacturer:
            match = re.search(r'Ryzen\s+\d\s+(\d)\d{3}', name)
            if match: generation = f"Serie Ryzen {match.group(1)}000"
        return {"Procesador": name, "Generación / Serie": generation, "Fabricante": manufacturer, "Núcleos Físicos": processor.NumberOfCores, "Procesadores Lógicos": processor.NumberOfLogicalProcessors, "Velocidad Máxima": f"{processor.MaxClockSpeed} MHz"}
    except Exception as e: return {"Error": f"No se pudo obtener la info del procesador: {e}"}

def get_installed_software(c):
    software_list = []
    try:
    
        for product in c.Win32_Product():
            software_list.append({"Nombre": product.Name, "Versión": product.Version, "Vendedor": product.Vendor})
        return sorted(software_list, key=lambda x: x['Nombre'])
    except Exception as e: return [{"Error": f"No se pudo obtener la lista de programas: {e}"}]

def get_installed_printers(c):
    printers = []
    try:
        for printer in c.Win32_Printer():
            printers.append({"Nombre": printer.Name, "Controlador": printer.DriverName, "Puerto": printer.PortName, "Default": "Sí" if printer.Default else "No"})
        return printers
    except Exception as e: return [{"Error": f"No se pudo obtener la lista de impresoras: {e}"}]