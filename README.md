# Excel Block Sorter

Windows icin hazirlanmis yerel bir Excel siralama aracidir. Finans veya
raporlama dosyalarinda toplam satirlarini ve bu toplamlarin altindaki detay
satirlarini otomatik olarak siralar.

## Ne Yapar?

Arac, Excel dosyasinda toplam satirlarini bulur ve iki seviyeli siralama yapar:

1. Toplam bloklarini buyukten kucuge siralar.
2. Her toplam blogunun altindaki detay satirlarini kendi icinde buyukten kucuge
   siralar.

Ornek:

```text
367.349.746 PERSONEL...
117.173.808 DISARIDAN...
60.077.177  NAKLIYE...
124.231.547 AMORTISMAN...
```

Cikti:

```text
367.349.746 PERSONEL...
124.231.547 AMORTISMAN...
117.173.808 DISARIDAN...
60.077.177  NAKLIYE...
```

`Genel Toplam` satiri varsa kategori gibi yukari tasinmaz, altta kalir.

## Ozellikler

- `.xlsx` dosyalarini isler.
- Tek dosya, coklu dosya veya klasor secilebilir.
- Sadece secilen kolona gore siralama yapar.
- Coklu siralama kolonu destekler: `J,K` gibi.
- Satirin sadece siralama hucrelerini degil, tum tablo satirini birlikte tasir.
- Orijinal Excel dosyasini degistirmez.
- Ciktilari masaustune veya secilen klasore kaydeder.
- Tek dosya icin yeni dosya adi verilebilir.

## Kurulum

Python 3.10+ onerilir.

Bagimliliklari yuklemek icin:

```powershell
pip install -r requirements.txt
```

Windows Python kurulumunda Tkinter genelde hazir gelir. GUI acilmazsa Python
kurulumunda Tkinter bileseni kontrol edilmelidir.

## GUI Ile Kullanma

GUI'yi baslatmak icin:

```powershell
Excel_Blok_Siralama.bat
```

Ardindan:

1. `Dosya ekle` veya `Klasor ekle` ile Excel dosyalarini secin.
2. `Siralama kolonu` alanina kolon yazin. Ornek: `J`.
3. Gerekirse coklu kolon yazin. Ornek: `J,K`.
4. Cikti klasorunu secin veya `Masaustu` dugmesini kullanin.
5. Tek dosya icin isterseniz `Yeni dosya adi` yazin.
6. `Calistir` dugmesine basin.

Varsayilan cikti klasoru:

```text
C:\Users\<kullanici>\Desktop\Excel_Sirali_Ciktilar
```

## Komut Satirindan Kullanma

Tek dosya:

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

Klasordeki tum Excel dosyalari:

```powershell
python excel_block_sorter.py --input . --output siralanmis --sort-column J
```

Sadece belirli sayfa:

```powershell
python excel_block_sorter.py --input Book1.xlsx --output siralanmis --sort-column J --sheet xxx
```

Dosya yazmadan kontrol:

```powershell
python excel_block_sorter.py --input Book1.xlsx --output siralanmis --sort-column J --dry-run
```

## Nasil Calisir?

Toplam satiri su sekilde bulunur:

- Secilen ana siralama kolonunda sayisal deger olmalidir.
- Ayni satirda dolgu rengi veya cift cizgi bicimi olmalidir.
- `2026 Fiili` gibi basliklar sayi kabul edilmez.

Blok mantigi:

- Her toplam satiri yeni bir blogun baslangicidir.
- Bir sonraki toplam satirina kadar olan satirlar o blogun detayidir.
- Bloklar toplam degerine gore siralanir.
- Detay satirlari kendi blogu icinde siralanir.

## Proje Dosyalari

```text
Excel_Blok_Siralama.bat     GUI'yi baslatir
excel_block_sorter_gui.py   Tkinter GUI
excel_block_sorter.py       Ana is motoru ve CLI
sort_excel_blocks.py        Eski komut adiyla uyumlu giris noktasi
test_excel_block_sorter.py  Testler
requirements.txt            Python bagimliliklari
README.md                   Proje dokumani
```

## Test

Testleri calistirmak icin:

```powershell
$env:PYTEST_DISABLE_PLUGIN_AUTOLOAD='1'
pytest -q
```

Derleme kontrolu:

```powershell
python -m py_compile excel_block_sorter.py excel_block_sorter_gui.py sort_excel_blocks.py
```

## Notlar

- `.xls` dosyalari desteklenmez, `.xlsx` kullanilmalidir.
- Orijinal Excel dosyalari degistirilmez.
- Excel dosyasi acik veya kilitliyse cikti kaydederken hata alinabilir.
- Siralanacak detay araliginda birlesik hucre varsa program guvenlik amaciyla
  ilgili dosyayi atlayabilir.
