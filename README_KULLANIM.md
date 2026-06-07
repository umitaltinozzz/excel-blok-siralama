# Excel Blok Siralama Araci

Bu proje, finans/raporlama Excel dosyalarinda kategori toplam satirlarini ve
alt detay satirlarini otomatik siralamak icin hazirlanmis yerel Windows
aracidir.

## Amac

Excel raporlarinda su yapi hedeflenir:

- Bir toplam/kategori satiri vardir. Ornek: `367.349.746 PERSONEL...`
- Bu satirin altinda o kategoriye ait detay satirlari vardir.
- Sonra bir sonraki toplam/kategori satiri gelir. Ornek: `117.173.808 ...`

Arac su islemleri yapar:

1. Toplam/kategori bloklarini secilen siralama kolonundaki toplam degere gore
   buyukten kucuge siralar.
2. Her toplam blogunun altindaki detay satirlarini da ayni kolona gore
   buyukten kucuge siralar.
3. `Genel Toplam` satirini kategori gibi yukari tasimaz; altta birakir.
4. Sadece siralama kolonunu degil, satirdaki tum tablo hucrelerini birlikte
   tasir. Boylece satir iliskileri bozulmaz.
5. Orijinal dosyayi degistirmez; cikti klasorune yeni `.xlsx` dosyasi yazar.

## Kullanicilar Icin GUI

GUI baslatma:

```powershell
Excel_Blok_Siralama.bat
```

Ana alanlar:

- `Dosya ekle`: Tek veya birden fazla Excel dosyasi secer.
- `Klasor ekle`: Bir klasordeki `.xlsx` dosyalarini isler.
- `Siralama kolonu`: Siralama kriteridir. Ornek: `J`.
- Coklu kriter ornegi: `J,K`.
  Once `J`, esitlikte `K` kullanilir.
- `Sayfa adi`: Bos birakilirsa tum sayfalar islenir.
- `Yeni dosya adi`: Sadece tek Excel dosyasi secildiginde kullanilir.
- `Coklu dosya eki`: Klasor veya coklu dosya seciminde dosya adina eklenir.
- `Masaustu`: Cikti klasorunu masaustundeki `Excel_Sirali_Ciktilar` klasorune
  ayarlar.

Varsayilan cikti klasoru:

```text
C:\Users\<kullanici>\Desktop\Excel_Sirali_Ciktilar
```

## Komut Satiri Kullanimlari

Tek kolon:

```powershell
python excel_block_sorter.py --input Book1.xlsx --output siralanmis --sort-column J
```

Coklu kolon:

```powershell
python excel_block_sorter.py --input Book1.xlsx --output siralanmis --sort-column J,K
```

Tek dosyada yeni ad:

```powershell
python excel_block_sorter.py --input Book1.xlsx --output siralanmis --sort-column J --output-name rapor_sirali.xlsx
```

Klasordeki tum `.xlsx` dosyalari:

```powershell
python excel_block_sorter.py --input . --output siralanmis --sort-column J
```

Sadece belirli sayfa:

```powershell
python excel_block_sorter.py --input Book1.xlsx --output siralanmis --sort-column J --sheet xxx
```

Kuru kosu:

```powershell
python excel_block_sorter.py --input Book1.xlsx --output siralanmis --sort-column J --dry-run
```

## Is Kurali Detayi

Toplam satiri tespiti:

- Secilen ana siralama kolonunda sayisal deger olmali.
- Ayni satirda dolgu rengi veya cift cizgi bicimi olmali.
- `2026 Fiili` gibi basliklar sayi kabul edilmez.

Blok tespiti:

- Her toplam satiri bir blogun baslangicidir.
- Bir sonraki toplam satirina kadar olan satirlar o blogun detayidir.

Blok siralama:

- Bloklar toplam satirindaki ana siralama kolonuna gore siralanir.
- `Genel Toplam` veya `Grand Total` etiketi iceren satir sabit tutulur.

Detay siralama:

- Her blogun detaylari secilen kolonlara gore siralanir.
- Bos veya sayi olmayan detay degerleri blogun alt tarafinda kalir.

Satir tasima:

- Program son kullanilan tablo kolonunu otomatik bulur.
- A kolonundan son dolu kolona kadar olan hucreler birlikte tasinir.

## Dosya Yapisi

```text
Excel_Blok_Siralama.bat   GUI'yi cift tikla baslatmak icin
excel_block_sorter_gui.py GUI katmani, Tkinter
excel_block_sorter.py     Ana is motoru ve CLI
sort_excel_blocks.py      Geriye donuk basit giris noktasi
test_excel_block_sorter.py Otomatik testler
requirements.txt          Python bagimliliklari
README_KULLANIM.md        Bu dokuman
```

## Kurulum

Python gereklidir. Test edilen ortam:

- Python 3.12
- openpyxl 3.1.x
- pytest 9.x

Bagimlilik kurulumu:

```powershell
pip install -r requirements.txt
```

Tkinter genelde Windows Python kurulumuyla gelir. GUI acilmazsa Python
kurulumunda Tkinter bileseni kontrol edilmelidir.

## Test

Pytest eklenti cakismalarini engellemek icin:

```powershell
$env:PYTEST_DISABLE_PLUGIN_AUTOLOAD='1'
pytest -q
```

Mevcut dogrulama durumu:

```text
12 passed
```

Derleme kontrolu:

```powershell
python -m py_compile excel_block_sorter.py excel_block_sorter_gui.py sort_excel_blocks.py
```

## Sinirlar ve Dikkat Edilecekler

- `.xlsx` desteklenir. `.xls` desteklenmez.
- Siralanacak detay satiri araliginda birlesik hucre varsa ilgili dosya/blok
  guvenlik nedeniyle atlanabilir.
- Formuller hucre degeri olarak tasinir; Excel'in formulleri yeniden hesaplamasi
  gerekebilir.
- Orijinal dosyalar degistirilmez.
- Excel dosyasi acik ve kilitliyse kaydetme sirasinda hata alinabilir.

## Beklenen Ornek Davranis

Ornek toplam sirasi:

```text
367.349.746 PERSONEL...
117.173.808 DISARIDAN...
60.077.177  NAKLIYE...
124.231.547 AMORTISMAN...
```

Cikti sonrasi:

```text
367.349.746 PERSONEL...
124.231.547 AMORTISMAN...
117.173.808 DISARIDAN...
60.077.177  NAKLIYE...
```

Her blogun alt detay satirlari da ayni siralama kolonuna gore buyukten kucuge
dizilir.
