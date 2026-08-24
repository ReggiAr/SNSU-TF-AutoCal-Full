import pyvisa

def scan_visa_devices():
    try:
        rm = pyvisa.ResourceManager()
    except Exception:
        try:
            rm = pyvisa.ResourceManager('@py')
        except Exception as e:
            print(f"Error: Gagal memuat VISA Backend ({e})")
            return

    print("=== MEMINDAI PERANGKAT VISA TERHUBUNG ===")
    try:
        perangkat_list = rm.list_resources()
    except Exception as e:
        print(f"Gagal mengambil daftar resources: {e}")
        return

    if not perangkat_list:
        print("Tidak ada instrumen atau port yang terdeteksi.")
        return

    print(f"Ditemukan {len(perangkat_list)} resource:\n")
    for addr in perangkat_list:
        print(f"Alamat : {addr}")
        try:
            inst = rm.open_resource(addr, timeout=1500)
            # Coba kirim query identifikasi standar
            respon_idn = inst.query("*IDN?").strip()
            print(f"Respon : {respon_idn}\n")
            inst.close()
        except Exception as e:
            print(f"Respon : Gagal membaca *IDN? ({e})\n")

if __name__ == "__main__":
    scan_visa_devices()