from ctypes import *
import sys
import time
import pyodbc
import threading
from datetime import datetime, timedelta

# --- AYARLAR ---
server = '172.30.134.15'
database = 'VALFSAN604'
username = 'pythonreporter'
password = '1212casecase,,'

GUN_SAYISI = 1
SESSISIZLIK_LIMITI = 10  # Cihazdan X saniye veri gelmezse bağlantıyı kes

pyodbc.pooling = False


# --- SQL'den cihaz listesi çek ---
def cihaz_listesini_sql_den_cek():
    """
    VLFPACSDEVICE tablosundan aktif cihazları çekip
    [{"ip": "...", "user": "...", "pass": "..."}] formatında döndürür.
    """
    cihazlar = []
    try:
        conn = pyodbc.connect(
            f"DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}",
            autocommit=True
        )
        cursor = conn.cursor()

        cursor.execute("""
            SELECT IPADRESS, USERNAME, PASSWORD
            FROM VLFPACSDEVICE
            WHERE ISACTIVE = 1
        """)

        rows = cursor.fetchall()
        for ip, user, pwd in rows:
            ip = (ip or "").strip()
            user = (user or "").strip()
            pwd = (pwd or "").strip()

            # boş kayıtları filtrele
            if ip and user and pwd:
                cihazlar.append({"ip": ip, "user": user, "pass": pwd})

        conn.close()

    except Exception as e:
        print(f"[SQL] Cihaz listesi çekilemedi: {e}")

    return cihazlar


# --- SQL FONKSİYONU ---
def sql_yaz(kart_no, zaman_str, cihaz_ip, seri_no):
    try:
        conn = pyodbc.connect(
            f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'
        )
        cursor = conn.cursor()

        check_query = "SELECT COUNT(*) FROM CihazLoglari WHERE KartNo=? AND TarihSaat=?"
        cursor.execute(check_query, (kart_no, zaman_str))

        if cursor.fetchone()[0] == 0:
            insert_query = """
                INSERT INTO CihazLoglari (TarihSaat, KartNo, Olay, CihazIP, SeriNo) 
                VALUES (?, ?, ?, ?, ?)
            """
            cursor.execute(insert_query, (zaman_str, kart_no, "Gecmis Kayit", cihaz_ip, seri_no))
            conn.commit()
            print(f"-> [SQL - {cihaz_ip}] Eklendi: {kart_no}")

        conn.close()
    except Exception as e:
        print(f"SQL Hatası ({cihaz_ip}): {e}")


def sql_prosedur_calistir():
    try:
        conn = pyodbc.connect(
            f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'
        )
        cursor = conn.cursor()

        cursor.execute("EXEC [dbo].[PDKSDATACOLLECT]")
        conn.commit()

        conn.close()
        print("-> [SQL] dbo.PDKSDATACOLLECT çalıştırıldı.")
    except Exception as e:
        print(f"SQL Prosedür Hatası: {e}")


# --- SDK YÜKLEME ---
if sys.platform == 'win32':
    sdk = CDLL(r".\HCNetSDK.dll")
else:
    sdk = cdll.LoadLibrary("./libhcnetsdk.so")

sdk.NET_DVR_Init()
sdk.NET_DVR_SetConnectTime(2000, 1)


# --- YAPILAR ---
class NET_DVR_TIME(Structure):
    _fields_ = [
        ("dwYear", c_int), ("dwMonth", c_int), ("dwDay", c_int),
        ("dwHour", c_int), ("dwMinute", c_int), ("dwSecond", c_int)
    ]


class NET_DVR_ACS_EVENT_COND(Structure):
    _fields_ = [
        ("dwSize", c_int), ("dwMajor", c_int), ("dwMinor", c_int),
        ("struStartTime", NET_DVR_TIME), ("struEndTime", NET_DVR_TIME),
        ("byCardNo", c_byte * 32), ("byName", c_byte * 32),
        ("byPicEnable", c_byte), ("byRes2", c_byte * 3),
        ("dwBeginSerialNo", c_int), ("dwEndSerialNo", c_int),
        ("byRes", c_byte * 244)
    ]


class NET_DVR_ACS_EVENT_CFG_SHORT(Structure):
    _fields_ = [
        ("dwSize", c_int), ("dwMajor", c_int), ("dwMinor", c_int),
        ("struTime", NET_DVR_TIME)
    ]


class NET_DVR_DEVICEINFO_V30(Structure):
    _fields_ = [
        ("sSerialNumber", c_byte * 48),
        ("byRes", c_byte * 200)
    ]


def py_time_to_struct(dt):
    t = NET_DVR_TIME()
    t.dwYear, t.dwMonth, t.dwDay = dt.year, dt.month, dt.day
    t.dwHour, t.dwMinute, t.dwSecond = dt.hour, dt.minute, dt.second
    return t


# --- CALLBACK FABRİKASI ---
CB_FUNC_TYPE = CFUNCTYPE(c_bool, c_int, c_void_p, c_int, c_void_p)


def callback_olustur(cihaz_ip, seri_no, durum_takip):
    def search_callback(dwType, pBuffer, dwBufLen, pUser):
        # Her veri geldiğinde zamanı güncelle
        durum_takip['son_aktivite'] = time.time()

        if dwType == 2 and pBuffer and dwBufLen > 0:
            try:
                # Zaman
                cfg = cast(pBuffer, POINTER(NET_DVR_ACS_EVENT_CFG_SHORT)).contents
                t = cfg.struTime
                zaman_str = f"{t.dwYear}-{t.dwMonth:02d}-{t.dwDay:02d} {t.dwHour:02d}:{t.dwMinute:02d}:{t.dwSecond:02d}"

                # Kart (offset 204)
                raw_data = string_at(pBuffer, dwBufLen)
                offset = 204
                if dwBufLen > offset:
                    raw_card = raw_data[offset: offset + 32]
                    kart_no = raw_card.partition(b'\0')[0].decode('utf-8', 'ignore').strip()

                    if kart_no and len(kart_no) > 2:
                        sql_yaz(kart_no, zaman_str, cihaz_ip, seri_no)

            except Exception as e:
                print(f"Hata ({cihaz_ip}): {e}")

        return True

    return CB_FUNC_TYPE(search_callback)


# --- CİHAZ GÖREVİ (THREAD) ---
def cihaz_gorevi(ip, user, password_):
    print(f"[{ip}] Bağlanılıyor...")

    device_info = NET_DVR_DEVICEINFO_V30()
    user_id = sdk.NET_DVR_Login_V30(
        ip.encode('utf-8'), 8000,
        user.encode('utf-8'), password_.encode('utf-8'),
        byref(device_info)
    )

    if user_id < 0:
        print(f"[{ip}] BAŞARISIZ! Login Hatası.")
        return

    try:
        seri_no = bytearray(device_info.sSerialNumber).partition(b'\0')[0].decode('utf-8', 'ignore')
    except:
        seri_no = "Bilinmiyor"

    takip_objesi = {'son_aktivite': time.time()}
    ozel_callback = callback_olustur(ip, seri_no, takip_objesi)

    search_cond = NET_DVR_ACS_EVENT_COND()
    search_cond.dwSize = sizeof(NET_DVR_ACS_EVENT_COND)
    search_cond.dwMajor = 0
    search_cond.dwMinor = 0

    now = datetime.now()
    start_dt = now - timedelta(days=GUN_SAYISI)
    search_cond.struStartTime = py_time_to_struct(start_dt)
    search_cond.struEndTime = py_time_to_struct(now)

    handle = sdk.NET_DVR_StartRemoteConfig(
        user_id, 2514, byref(search_cond), sizeof(search_cond), ozel_callback, None
    )

    if handle < 0:
        print(f"[{ip}] Sorgu Başlamadı.")
        sdk.NET_DVR_Logout(user_id)
        return

    print(f"[{ip}] Veri çekiliyor... (Sessizlik limiti: {SESSISIZLIK_LIMITI}sn)")

    try:
        while True:
            time.sleep(0.5)
            gecen_sure = time.time() - takip_objesi['son_aktivite']
            if gecen_sure > SESSISIZLIK_LIMITI:
                print(f"[{ip}] Veri akışı tamamlandı (Zaman aşımı).")
                break

    except KeyboardInterrupt:
        pass

    sdk.NET_DVR_StopRemoteConfig(handle)
    sdk.NET_DVR_Logout(user_id)
    print(f"[{ip}] Bağlantı kapatıldı.")


# --- ANA PROGRAM ---
def main():
    cihaz_listesi = cihaz_listesini_sql_den_cek()

    if not cihaz_listesi:
        print("Aktif cihaz bulunamadı (VLFPACSDEVICE.ISACTIVE=1). Program durduruldu.")
        sdk.NET_DVR_Cleanup()
        return

    print(f"Toplam {len(cihaz_listesi)} cihaz taranacak.")
    threads = []

    for cihaz in cihaz_listesi:
        t = threading.Thread(
            target=cihaz_gorevi,
            args=(cihaz["ip"], cihaz["user"], cihaz["pass"])
        )
        threads.append(t)
        t.start()

    for t in threads:
        t.join()

    sql_prosedur_calistir()

    print("-" * 30)
    print("TÜM CİHAZLARIN İŞLEMİ BİTTİ.")
    print("-" * 30)
    sdk.NET_DVR_Cleanup()

if __name__ == "__main__":
    main()
