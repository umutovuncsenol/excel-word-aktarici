# Excel Word Aktarıcı

Excel dosyasındaki yedi tabloyu bulur ve karşılık gelen Word tablolarına aktarır.
Aktarım başlamadan önce Excel ve Word tablo/sütun eşleştirmelerini gösterir; bütün
seçimler kullanıcı tarafından değiştirilebilir.

## Windows kullanımı

1. Releases bölümündeki `ExcelWordAktarici.exe` dosyasını indirin.
2. Programı açın.
3. Excel dosyasını, Word şablonunu ve çıktı konumunu seçin.
4. Otomatik bulunan tablo ve sütun eşleştirmelerini kontrol edin.
5. `Onayla ve Aktar` düğmesine basın.

Program ağ bağlantısı kullanmaz. Seçilen belgeler yalnızca bilgisayarınızda işlenir.

## Kaynak koddan çalıştırma

```powershell
pip install -r requirements.txt
python excel_word_aktarici.py
```
