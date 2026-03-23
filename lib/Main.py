from ctypes import *
import sys
import time
import pyodbc
import threading
from datetime import datetime, timedelta

import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# -------------------- SQL BAĞLANTI AYARLARI --------------------
server = '172.30.134.15'
database = 'VALFSAN604'
username = 'pythonreporter'
password = '1212casecase,,'

SESSISIZLIK_LIMITI = 10  # Cihazdan X saniye veri gelmezse bağlantıyı kes
pyodbc.pooling = False

# -------------------- MAIL AYARLARI (Sizdeki örneğe benzer) --------------------
MAIL_MODE = "relay"   # "relay" veya "o365"

sender_email = "canias@valfsan.com.tr"
MAIL_TO = "canias@valfsan.com.tr"
MAIL_CC = ""

SMTP_DEBUG = True
SMTP_VERIFY_CERT = False
SMTP_CA_FILE = ""
SMTP_SLEEP_AFTER_STARTTLS = 0

if MAIL_MODE == "o365":
    mail_server = "smtp.office365.com"
    mail_port = 587
    SMTP_USE_STARTTLS = True
    SMTP_REQUIRE_AUTH = True
    SMTP_USERNAME = "canias@valfsan.com.tr"
    SMTP_PASSWORD = ""
else:
    # Yerel Exchange relay
    mail_server = "mail.valfsan.com.tr"
    mail_port = 25
    SMTP_USE_STARTTLS = False
    SMTP_REQUIRE_AUTH = False
    SMTP_USERNAME = ""
    SMTP_PASSWORD = ""

##

# -------------------- SQL'DEN CİHAZ LİSTESİ ÇEKME --------------------
def cihazlari_sql_den_getir():
    """
    VLFPACSDEVICE tablosundan aktif cihazları çeker.
    DDAY her cihazın kaç gün geriye taranacağını belirler.
    NAME mail raporunda kullanılacak.
    """
    conn = None
    try:
        conn = pyodbc.connect(
            f"DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}"
        )
        cursor = conn.cursor()

        query = """
            SELECT IPADRESS, USERNAME, PASSWORD, DDAY, NAME
            FROM VLFPACSDEVICE
            WHERE ISACTIVE = 1
        """
        cursor.execute(query)

        cihaz_listesi = []
        for row in cursor.fetchall():
            ip = (row.IPADRESS or "").strip()
            user_ = (row.USERNAME or "").strip()
            pass_ = (row.PASSWORD or "").strip()
            name_ = (row.NAME or "").strip()

            try:
                dday = int(row.DDAY) if row.DDAY is not None else 1
            except:
                dday = 1

            if dday <= 0:
                dday = 1

            if ip and user_ and pass_:
                cihaz_listesi.append({
                    "ip": ip,
                    "user": user_,
                    "pass": pass_,
                    "dday": dday,
                    "name": name_
                })

        return cihaz_listesi

    except Exception as e:
        print(f"CIHAZ_LISTESI SQL okuma hatası: {e}")
        return []
    finally:
        if conn:
            conn.close()


# -------------------- SQL YAZMA --------------------
def sql_yaz(kart_no, zaman_str, cihaz_ip, seri_no):
    """
    Kayıt eklendiyse True, eklenmediyse False döner.
    """
    conn = None
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
            return True

        return False
    except Exception as e:
        print(f"SQL Hatası ({cihaz_ip}): {e}")
        return False
    finally:
        if conn:
            conn.close()


def sql_prosedur_calistir():
    conn = None
    try:
        conn = pyodbc.connect(
            f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'
        )
        cursor = conn.cursor()

        cursor.execute("EXEC [dbo].[PDKSDATACOLLECT]")
        conn.commit()

        print("-> [SQL] dbo.PDKSDATACOLLECT çalıştırıldı.")
    except Exception as e:
        print(f"SQL Prosedür Hatası: {e}")
    finally:
        if conn:
            conn.close()



def mail_gonder_exchange(cihaz_sonuclari):
    # NO_DATA da mail tetiklesin
    hatali = [
        r for r in cihaz_sonuclari
        if r.get("status") in ("LOGIN_FAIL", "QUERY_FAIL", "EXCEPTION")
    ]

    if not hatali:
        print("-> [MAIL] Hatalı cihaz yok, mail atılmadı.")
        return

    now_str = time.strftime("%d-%m-%Y %H:%M:%S")

    def row_html(r):
        return (
            f"<li><b>{r.get('name','')}</b> | IP: {r.get('ip')} | "
            f"Eklenen: {r.get('inserted',0)} | Durum: {r.get('status')} | "
            f"Sebep: {r.get('reason','')}</li>"
        )

    hatali_html = "\n".join([row_html(r) for r in hatali])
    tum_html = "\n".join([row_html(r) for r in cihaz_sonuclari])

    body = f"""
    <html>
      <body>
        <p>Merhaba,</p>
        <p><b>{now_str}</b> itibarıyla PDKS cihaz taramasında
        <b>veri çekilemeyen / hatalı cihaz(lar)</b> tespit edildi.</p>

        <p><b>Problemli cihazlar:</b></p>
        <ul>
          {hatali_html}
        </ul>

        <p><b>Tüm cihazlar özeti:</b></p>
        <ul>
          {tum_html}
        </ul>

        <p>Saygılarımızla,</p>
      </body>
    </html>
    """

    msg = MIMEMultipart()
    msg["From"] = sender_email
    msg["To"] = MAIL_TO
    if MAIL_CC.strip():
        msg["Cc"] = MAIL_CC
    msg["Subject"] = f"PDKS - Cihaz Raporu ({now_str})"
    msg.attach(MIMEText(body, "html", "utf-8"))

    recipients = [x.strip() for x in MAIL_TO.replace(";", ",").split(",") if x.strip()]
    if MAIL_CC.strip():
        recipients += [x.strip() for x in MAIL_CC.replace(";", ",").split(",") if x.strip()]

    try:
        if SMTP_VERIFY_CERT:
            context = ssl.create_default_context(cafile=SMTP_CA_FILE) if SMTP_CA_FILE else ssl.create_default_context()
        else:
            context = ssl.create_default_context()
            context.check_hostname = False
            context.verify_mode = ssl.CERT_NONE

        with smtplib.SMTP(mail_server, mail_port, timeout=60) as smtp:
            if SMTP_DEBUG:
                smtp.set_debuglevel(1)

            smtp.ehlo()

            # STARTTLS sadece gerçekten destekleniyorsa çalıştır
            if SMTP_USE_STARTTLS:
                if not smtp.has_extn("starttls"):
                    raise RuntimeError(
                        f"SMTP sunucusu STARTTLS desteklemiyor: {mail_server}:{mail_port}"
                    )
                smtp.starttls(context=context)
                smtp.ehlo()

                if SMTP_SLEEP_AFTER_STARTTLS:
                    time.sleep(SMTP_SLEEP_AFTER_STARTTLS)

            if SMTP_REQUIRE_AUTH:
                smtp.login(SMTP_USERNAME, SMTP_PASSWORD)

            smtp.sendmail(sender_email, recipients, msg.as_string())
            print("-> [MAIL] Email successfully sent!")

    except Exception as e:
        print(f"[MAIL HATASI] Failed to send email: {e}")

# -------------------- SDK YÜKLEME --------------------
if sys.platform == 'win32':
    sdk = CDLL(r".\HCNetSDK.dll")
else:
    sdk = cdll.LoadLibrary("./libhcnetsdk.so")

sdk.NET_DVR_Init()
sdk.NET_DVR_SetConnectTime(2000, 1)


# -------------------- YAPILAR --------------------
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


# -------------------- CALLBACK --------------------
CB_FUNC_TYPE = CFUNCTYPE(c_bool, c_int, c_void_p, c_int, c_void_p)


def callback_olustur(cihaz_ip, seri_no, durum_takip):
    def search_callback(dwType, pBuffer, dwBufLen, pUser):
        durum_takip["son_aktivite"] = time.time()

        if dwType == 2 and pBuffer and dwBufLen > 0:
            try:
                cfg = cast(pBuffer, POINTER(NET_DVR_ACS_EVENT_CFG_SHORT)).contents
                t = cfg.struTime
                zaman_str = f"{t.dwYear}-{t.dwMonth:02d}-{t.dwDay:02d} {t.dwHour:02d}:{t.dwMinute:02d}:{t.dwSecond:02d}"

                raw_data = string_at(pBuffer, dwBufLen)
                offset = 204
                if dwBufLen > offset:
                    raw_card = raw_data[offset: offset + 32]
                    kart_no = raw_card.partition(b'\0')[0].decode('utf-8', 'ignore').strip()

                    if kart_no and len(kart_no) > 2:
                        eklendi = sql_yaz(kart_no, zaman_str, cihaz_ip, seri_no)
                        if eklendi:
                            durum_takip["inserted_count"] += 1

            except Exception as e:
                print(f"Hata ({cihaz_ip}): {e}")

        return True

    return CB_FUNC_TYPE(search_callback)


# -------------------- CİHAZ GÖREVİ (THREAD) --------------------
def cihaz_gorevi(ip, user, password_, gun_sayisi, cihaz_adi, sonuc_listesi, sonuc_lock):
    print(f"[{ip}] Bağlanılıyor... (DDAY={gun_sayisi})")

    result = {
        "ip": ip,
        "name": cihaz_adi,
        "dday": gun_sayisi,
        "seri_no": "",
        "inserted": 0,
        "status": "UNKNOWN",
        "reason": ""
    }

    device_info = NET_DVR_DEVICEINFO_V30()
    user_id = sdk.NET_DVR_Login_V30(
        ip.encode("utf-8"), 8000,
        user.encode("utf-8"), password_.encode("utf-8"),
        byref(device_info)
    )

    if user_id < 0:
        print(f"[{ip}] BAŞARISIZ! Login Hatası.")
        result["status"] = "LOGIN_FAIL"
        result["reason"] = "Login başarısız"
        with sonuc_lock:
            sonuc_listesi.append(result)
        return

    try:
        seri_no = bytearray(device_info.sSerialNumber).partition(b'\0')[0].decode('utf-8', 'ignore')
    except:
        seri_no = "Bilinmiyor"

    result["seri_no"] = seri_no

    takip_objesi = {"son_aktivite": time.time(), "inserted_count": 0}
    ozel_callback = callback_olustur(ip, seri_no, takip_objesi)

    search_cond = NET_DVR_ACS_EVENT_COND()
    search_cond.dwSize = sizeof(NET_DVR_ACS_EVENT_COND)
    search_cond.dwMajor = 0
    search_cond.dwMinor = 0

    now = datetime.now()
    start_dt = now - timedelta(days=gun_sayisi)
    search_cond.struStartTime = py_time_to_struct(start_dt)
    search_cond.struEndTime = py_time_to_struct(now)

    handle = -1
    try:
        handle = sdk.NET_DVR_StartRemoteConfig(
            user_id, 2514, byref(search_cond), sizeof(search_cond), ozel_callback, None
        )

        if handle < 0:
            print(f"[{ip}] Sorgu Başlamadı.")
            result["status"] = "QUERY_FAIL"
            result["reason"] = "NET_DVR_StartRemoteConfig başarısız"
            return

        print(f"[{ip}] Veri çekiliyor... (Sessizlik limiti: {SESSISIZLIK_LIMITI}sn)")

        while True:
            time.sleep(0.5)
            if (time.time() - takip_objesi["son_aktivite"]) > SESSISIZLIK_LIMITI:
                print(f"[{ip}] Veri akışı tamamlandı (Zaman aşımı).")
                break

        result["inserted"] = takip_objesi["inserted_count"]
        result["status"] = "OK" if result["inserted"] > 0 else "NO_DATA"
        result["reason"] = "Başarılı" if result["inserted"] > 0 else f"{gun_sayisi} gün aralığında kayıt gelmedi"

    except Exception as e:
        print(f"[{ip}] Thread exception: {e}")
        result["status"] = "EXCEPTION"
        result["reason"] = str(e)

    finally:
        try:
            if handle is not None and handle >= 0:
                sdk.NET_DVR_StopRemoteConfig(handle)
        except:
            pass

        try:
            sdk.NET_DVR_Logout(user_id)
        except:
            pass

        print(f"[{ip}] Bağlantı kapatıldı.")
        with sonuc_lock:
            sonuc_listesi.append(result)


# -------------------- ANA PROGRAM --------------------
def main():
    CIHAZ_LISTESI = cihazlari_sql_den_getir()

    if not CIHAZ_LISTESI:
        print("Aktif cihaz bulunamadı. Program kapatılıyor.")
        sdk.NET_DVR_Cleanup()
        return

    print(f"Toplam {len(CIHAZ_LISTESI)} cihaz taranacak.")
    threads = []

    sonuc_listesi = []
    sonuc_lock = threading.Lock()

    for cihaz in CIHAZ_LISTESI:
        t = threading.Thread(
            target=cihaz_gorevi,
            args=(cihaz["ip"], cihaz["user"], cihaz["pass"], cihaz["dday"], cihaz.get("name", ""), sonuc_listesi, sonuc_lock)
        )
        threads.append(t)
        t.start()

    for t in threads:
        t.join()

    sql_prosedur_calistir()

    # Sadece veri çekilemeyen cihaz varsa mail atar
    mail_gonder_exchange(sonuc_listesi)

    print("-" * 30)
    print("TÜM CİHAZLARIN İŞLEMİ BİTTİ.")
    print("-" * 30)
    sdk.NET_DVR_Cleanup()


if __name__ == "__main__":
    main()
